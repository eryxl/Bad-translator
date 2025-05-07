function onOpen() {
  DocumentApp.getUi().createMenu('HyperTranslate')
    .addItem('Translate Selected Text 100 Times', 'hyperTranslate')
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
  const textSegments = [];

  // Gather all text segments
  elements.forEach(element => {
    const elem = element.getElement();
    if (elem.editAsText) {
      const textElem = elem.asText();
      const start = element.getStartOffset();
      const end = element.getEndOffsetInclusive();

      if (start !== undefined && end !== undefined) {
        try {
          const segment = textElem.getText().substring(start, end + 1);
          if (segment.trim()) {
            selectedText += segment + ' ';
            textSegments.push({
              element: textElem,
              start,
              end
            });
          }
        } catch (e) {
          // Skip problematic elements
        }
      }
    }
  });

  if (!selectedText.trim()) {
    DocumentApp.getUi().alert('No valid text found in selection.');
    return;
  }

  let translatedText = selectedText.trim();

  const supportedLanguages = [
    'af','ar','bg','bn','ca','cs','cy','da','de','el','en','eo','es','et','fa',
    'fi','fr','gu','he','hi','hr','ht','hu','id','is','it','ja','ka','ko','lt',
    'lv','mk','mr','ms','mt','nl','no','pl','pt','ro','ru','sk','sl','sq','sv',
    'sw','ta','te','th','tl','tr','uk','ur','vi','zh'
  ];

  for (let i = 0; i < 50; i++) {
    let randomLang = supportedLanguages[Math.floor(Math.random() * supportedLanguages.length)];
    if (randomLang === 'en') {
      i--;
      continue;
    }

    try {
      translatedText = LanguageApp.translate(translatedText, 'en', randomLang);
      translatedText = LanguageApp.translate(translatedText, randomLang, 'en');
    } catch (error) {
      DocumentApp.getUi().alert(`Error during translation ${i + 1}:\n\n${error.message}`);
      return;
    }
  }

  // Replace all selected segments with the translated text (distribute it evenly or entirely)
  const perSegmentText = translatedText; // Alternatively, divide if desired

  textSegments.forEach(segment => {
    const { element, start, end } = segment;
    element.deleteText(start, end);
    element.insertText(start, perSegmentText);
  });

  DocumentApp.getUi().alert('Translation complete and text replaced!');
}
