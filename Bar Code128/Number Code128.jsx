/*******************************************************************************
 * Code 128-C Barcode Generator for Adobe Illustrator
 * Generates numbers only
 * Created by Nikolay Rudenko
 ******************************************************************************/

// Определения паттернов для Code 128 (C)
var CODE_128C_PATTERNS = {
    startCode: "11010011100", // Start Code C (значение 105)
    stopCode: "1100011101011", // Stop Code (универсальный)
    patterns: [ // Паттерны для значений 00-102
        "11011001100", "11001101100", "11001100110", "10010011000",
        "10010001100", "10001001100", "10011001000", "10011000100",
        "10001100100", "11001001000", "11001000100", "11000100100",
        "10110011100", "10011011100", "10011001110", "10111001100",
        "10011101100", "10011100110", "11001110010", "11001011100",
        "11001001110", "11011100100", "11001110100", "11101101110",
        "11101001100", "11100101100", "11100100110", "11101100100",
        "11100110100", "11100110010", "11011011000", "11011000110",
        "11000110110", "10100011000", "10001011000", "10001000110",
        "10110001000", "10001101000", "10001100010", "11010001000",
        "11000101000", "11000100010", "10110111000", "10110001110",
        "10001101110", "10111011000", "10111000110", "10001110110",
        "11101110110", "11010001110", "11000101110", "11011101000",
        "11011100010", "11011101110", "11101011000", "11101000110",
        "11100010110", "11101101000", "11101100010", "11100011010",
        "11101111010", "11001000010", "11110001010", "10100110000",
        "10100001100", "10010110000", "10010000110", "10000101100",
        "10000100110", "10110010000", "10110000100", "10011010000",
        "10011000010", "10000110100", "10000110010", "11000010010",
        "11001010000", "11110111010", "11000010100", "10001111010",
        "10100111100", "10010111100", "10010011110", "10111100100",
        "10011110100", "10011110010", "11110100100", "11110010100",
        "11110010010", "11011011110", "11011110110", "11110110110",
        "10101111000", "10100011110", "10001011110", "10111101000",
        "10111100010", "11110101000", "11110100010", "10111011110",
        "10111101110", "11101011110", "11110101110"
    ]
};

/**
 * Главная функция, запускающая создание штрих-кодов для выделенных текстовых объектов.
 */
function createCode128BarcodesForSelection() {
    if (app.documents.length === 0) {
        alert("Пожалуйста, сначала откройте документ!");
        return;
    }

    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert("Пожалуйста, выделите один или несколько текстовых объектов!");
        return;
    }

    for (var i = 0; i < doc.selection.length; i++) {
        var selectedItem = doc.selection[i];
        if (selectedItem.typename === "TextFrame" || selectedItem.typename === "PointText") {
            createBarcodeForTextItem(selectedItem);
        } else {
            alert("Элемент " + (i + 1) + " не является текстовым объектом. Пропускаем.");
        }
    }
}

/**
 * Создает штрихкод для конкретного текстового элемента.
 * @param {TextFrame | PointText} textItem - Текстовый объект Illustrator.
 */
function createBarcodeForTextItem(textItem) {
    var message = textItem.contents;
    if (message.length === 0) {
        alert("Текстовый фрейм пуст. Пропускаем.");
        return;
    }

    // Генерируем паттерн Code 128-C
    var pattern = generateCode128CPattern(message);
    if (!pattern) { // Если генерация не удалась (из-за неверного ввода), прерываем
        return;
    }

    var matrix = textItem.matrix;
    var angle = -180 / Math.PI * Math.atan2(matrix.mValueC, matrix.mValueD);
    var textBounds;
    
    // Корректно получаем границы для повернутого текста
    if (angle !== 0) {
        var tempText = textItem.duplicate();
        tempText.rotate(-angle);
        textBounds = tempText.geometricBounds;
        tempText.remove(); // Очищаем временный объект
    } else {
        textBounds = textItem.geometricBounds;
    }

    var textWidth = textBounds[2] - textBounds[0];
    var textHeight = Math.abs(textBounds[1] - textBounds[3]);

    // Создаем группу для хранения текста и штрих-кода
    var group = app.activeDocument.groupItems.add();
    textItem.move(group, ElementPlacement.PLACEATEND);

    // --- Расчет размеров и отрисовка штрих-кода ---
    var targetBarcodeWidth = textWidth * 1.2;
    var barWidth = targetBarcodeWidth / pattern.length;
    var barcodeHeight = textHeight * 4.5;
    var barcodeWidth = pattern.length * barWidth;
    
    var startX = textBounds[0] + (textWidth - barcodeWidth) / 2;
    var currentX = startX;

    var barcodeGroup = group.groupItems.add();
    var blackColor = new CMYKColor();
    blackColor.black = 100;

    // Рисуем каждый штрих
    for (var i = 0; i < pattern.length; i++) {
        if (pattern[i] === "1") {
            var bar = barcodeGroup.pathItems.rectangle(
                textBounds[1] + 3, // Расположение над текстом с небольшим отступом
                currentX,
                barWidth,
                -barcodeHeight
            );
            bar.filled = true;
            bar.fillColor = blackColor;
            bar.stroked = false;
        }
        currentX += barWidth;
    }
    
    // --- Поворот штрих-кода ---
    // Определяем угол из матрицы трансформации текста
    var rotationAngle = -180 / Math.PI * Math.atan2(textItem.matrix.mValueC, textItem.matrix.mValueD);
    
    // Находим центр вращения (центр текстового объекта)
    var rotationX = textItem.left + textItem.width / 2;
    var rotationY = textItem.top - textItem.height / 2;

    // Применяем поворот к группе со штрихкодом
    rotate_around_point(barcodeGroup, rotationY, rotationX, rotationAngle);
}

/**
 * Генерирует строку из '1' и '0' для штрих-кода Code 128-C.
 * @param {string} message - Входная строка. Должна содержать четное количество цифр.
 * @returns {string} - Двоичный паттерн штрих-кода или пустая строка в случае ошибки.
 */
function generateCode128CPattern(message) {
    // 1. Проверка на корректность ввода
    if (!/^\d+$/.test(message)) {
        alert("Ошибка: Входные данные для Code 128-C могут содержать только цифры.\nОбнаружено: '" + message + "'");
        return "";
    }
    if (message.length % 2 !== 0) {
        alert("Ошибка: Входные данные для Code 128-C должны иметь четное количество цифр.\nДлина строки: " + message.length);
        return "";
    }

    // 2. Инициализация
    var pattern = CODE_128C_PATTERNS.startCode;
    var checksum = 105; // Начальное значение для Start Code C

    // 3. Кодирование пар цифр
    for (var i = 0; i < message.length; i += 2) {
        var pairStr = message.substring(i, i + 2);
        var value = parseInt(pairStr, 10);
        
        pattern += CODE_128C_PATTERNS.patterns[value];
        
        // Вес для контрольной суммы зависит от позиции ПАРЫ, а не символа
        var weight = (i / 2) + 1;
        checksum += weight * value;
    }

    // 4. Расчет и добавление контрольной суммы
    checksum = checksum % 103;
    pattern += CODE_128C_PATTERNS.patterns[checksum];

    // 5. Добавление стоп-кода
    pattern += CODE_128C_PATTERNS.stopCode;

    return pattern;
}

/**
 * Вспомогательная функция для поворота объекта вокруг произвольной точки.
 * @param {PageItem} obj - Объект для поворота.
 * @param {number} roto_X - Координата X центра вращения.
 * @param {number} roto_Y - Координата Y центра вращения.
 * @param {number} angle - Угол поворота в градусах.
 */
function rotate_around_point(obj, roto_X, roto_Y, angle) {
    // Смещаем к началу координат документа
    obj.translate(-roto_Y, -roto_X);
    // Вращаем
    obj.rotate(angle, true, true, true, true, Transformation.DOCUMENTORIGIN);
    // Смещаем обратно
    obj.translate(roto_Y, roto_X);
}

// Запуск скрипта с обработкой ошибок
try {
    createCode128BarcodesForSelection();
} catch (e) {
    alert("Ошибка при создании штрихкода: " + e.message);
}