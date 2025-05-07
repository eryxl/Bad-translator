function onOpen() {
  DocumentApp.getUi().createMenu('HyperTranslate')
    .addItem('Translate Selected Text 100 Times', 'hyperTranslate')
    .addToUi();
}

function hyperTranslate() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  // Check if there's any selection
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
      
      // Only proceed if we have valid offsets
      if (start !== undefined && end !== undefined) {
        const segment = textElem.getText().substring(start, end + 1);
        if (segment.trim()) {  // Only add non-empty segments
          selectedText += segment + '\n';
        }
      }
    }
  });

  // Check if we actually got any text
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

  // Perform translations
  for (let i = 0; i < 50; i++) {  // Reduced iterations for reliability
    let randomLang = supportedLanguages[Math.floor(Math.random() * supportedLanguages.length)];
    
    // Avoid translating to/from English until the final step
    if (randomLang === 'en') {
      i--;
      continue;
    }

    try {
      // First translate to random language
      translatedText = LanguageApp.translate(translatedText, 'en', randomLang);
      
      // Then translate back to English
      translatedText = LanguageApp.translate(translatedText, randomLang, 'en');
      
      // Show progress every 10 iterations
      if ((i + 1) % 10 === 0) {
        DocumentApp.getUi().alert(`Completed ${i + 1} translations...`);
      }
    } catch (error) {
      DocumentApp.getUi().alert(`Error during translation ${i + 1}:\n\n${error.message}`);
      return;
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
