function onOpen() {
  DocumentApp.getUi().createMenu('HyperTranslate')
      .addItem('Translate Selected Text 50 Times', 'hyperTranslate')
      .addToUi();
}

function hyperTranslate() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();

  if (!selection) {
    DocumentApp.getUi().alert('Please select some text to translate.');
    return;
  }

  const elements = selection.getRangeElements();
  let selectedText = '';

  // Gather selected text
  elements.forEach(element => {
    const textElem = element.getElement().asText();
    if (textElem) {
      const start = element.getStartOffset();
      const end = element.getEndOffsetInclusive();
      if (start !== undefined && end !== undefined) {
        selectedText += textElem.getText().substring(start, end + 1) + '\n';
      }
    }
  });

  if (!selectedText.trim()) {
    DocumentApp.getUi().alert('No valid text found in selection.');
    return;
  }

  let translatedText = selectedText.trim();

  const supportedLanguages = [
    'af','ar','bg','bn','ca','cs','cy','da','de','el','en','eo','es','et','fa','fi','fr','gu',
    'he','hi','hr','ht','hu','id','is','it','ja','ka','ko','lt','lv','mk','mr','ms','mt','nl',
    'no','pl','pt','ro','ru','sk','sl','sq','sv','sw','ta','te','th','tl','tr','uk','ur','vi','zh'
  ];

  // Perform 50 translations
  for (let i = 0; i < 50; i++) {
    let randomLang = supportedLanguages[Math.floor(Math.random() * supportedLanguages.length)];

    // Avoid translating to/from English until the final step
    if (randomLang === 'en') {
      i--;
      continue;
    }

    try {
      translatedText = LanguageApp.translate(translatedText, 'en', randomLang);
    } catch (error) {
      DocumentApp.getUi().alert('Error during translation:\n\n' + error.message);
      return;
    }

    // Translate back to English each time
    try {
      translatedText = LanguageApp.translate(translatedText, randomLang, 'en');
    } catch (error) {
      DocumentApp.getUi().alert('Error translating back to English:\n\n' + error.message);
      return;
    }

    if ((i + 1) % 10 === 0) {
      DocumentApp.getUi().alert(`Completed ${i + 1} translations...`);
    }
  }

  // Replace original text with final translated version
  elements.forEach(element => {
    const textElem = element.getElement().asText();
    if (textElem) {
      const start = element.getStartOffset();
      const end = element.getEndOffsetInclusive();
      if (start !== undefined && end !== undefined) {
        textElem.deleteText(start, end);
        textElem.insertText(start, translatedText);
      }
    }
  });

  DocumentApp.getUi().alert('Translation complete and text replaced!');
}
