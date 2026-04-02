/**********************************************************
Mykola Rudenko © Copyright 2025 All Rights Reserved 
*********************************************************/
// Скрипт для Adobe Illustrator - Lapki v02
// Замінює прямі подвійні лапки (" ") на друкарські лапки-ялинки (« »)

function replaceQuotesInTextFrames() {
    if (app.documents.length === 0) {
        alert("Відкрийте документ в Illustrator перед запуском скрипта.");
        return;
    }

    var doc = app.activeDocument;
    var textFrames = doc.textFrames;

    for (var i = 0; i < textFrames.length; i++) {
        var tf = textFrames[i];
        var chars = tf.textRange.characters;
        var quoteOpen = true;

        for (var j = 0; j < chars.length; j++) {
            var c = chars[j].contents;

            if (c === '"' || c === '“' || c === '”') {
                chars[j].contents = quoteOpen ? '«' : '»';
                quoteOpen = !quoteOpen;
            }
        }
    }

    alert("Заміна лапок відбулась!");
}

replaceQuotesInTextFrames();


