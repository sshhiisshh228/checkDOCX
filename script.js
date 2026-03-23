// Глобальные переменные
let currentResults = null;
let selectedParameters = new Set();

// Список всех доступных параметров
const allParameters = [
    // Параметры форматирования
    { id: 'alignment', name: 'Выравнивание', category: 'formatting' },
    { id: 'specialIndentation', name: 'Специальные отступы', category: 'formatting' },
    { id: 'keepWithNext', name: 'Связать со следующим', category: 'formatting' },
    { id: 'outlineLevel', name: 'Уровень структуры', category: 'formatting' },
    { id: 'keepLinesTogether', name: 'Не разрывать абзац', category: 'formatting' },
    { id: 'pageBreakBefore', name: 'Разрыв страницы', category: 'formatting' },
    
    // Параметры интервалов
    { id: 'spaceBefore', name: 'Отступ перед', category: 'spacing' },
    { id: 'spaceAfter', name: 'Отступ после', category: 'spacing' },
    { id: 'lineSpacing', name: 'Межстрочный интервал', category: 'spacing' },
    { id: 'leftIndentation', name: 'Левый отступ', category: 'spacing' },
    { id: 'rightIndentation', name: 'Правый отступ', category: 'spacing' },
    
    // Параметры текста
    { id: 'fullBold', name: 'Полностью жирный', category: 'text' },
    { id: 'fullItalic', name: 'Полностью курсив', category: 'text' },
    { id: 'isLowercase', name: 'Строчные буквы', category: 'text' },
    { id: 'isUppercase', name: 'Прописные буквы', category: 'text' },
    { id: 'specialSymbolsCount', name: 'Специальные символы', category: 'text' },
    
    // Параметры шрифтов
    { id: 'fonts', name: 'Шрифты', category: 'font' },
    { id: 'fontSizes', name: 'Размеры шрифта', category: 'font' },
    { id: 'isTimesNewRoman', name: 'Times New Roman', category: 'font' },
    { id: 'mainFontSize', name: 'Основной размер', category: 'font' },
    { id: 'fontSizeConsistency', name: 'Единый размер', category: 'font' },
    
    // Параметры класса
    { id: 'class', name: 'Класс параграфа', category: 'class' },
    { id: 'computed_mark', name: 'Вычисленная оценка', category: 'class' }
];

// Инициализация при загрузке страницы
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    fileInput.addEventListener('change', handleFileSelect);
    
    // Добавляем панель выбора параметров
    addParametersSelector();
    
    // Добавляем поддержку drag and drop
    const uploadArea = document.querySelector('.upload-area');
    
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#667eea';
        uploadArea.style.backgroundColor = 'rgba(102, 126, 234, 0.1)';
    });
    
    uploadArea.addEventListener('dragleave', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#e0e0e0';
        uploadArea.style.backgroundColor = 'transparent';
    });
    
    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.style.borderColor = '#e0e0e0';
        uploadArea.style.backgroundColor = 'transparent';
        
        const files = e.dataTransfer.files;
        if (files.length > 0 && files[0].name.toLowerCase().endsWith('.docx')) {
            fileInput.files = files;
            handleFileSelect({ target: { files: [files[0]] } });
        } else {
            showError('Пожалуйста, загрузите файл в формате DOCX');
        }
    });
});

function addParametersSelector() {
    const uploadArea = document.querySelector('.upload-area');
    
    const selector = document.createElement('div');
    selector.className = 'parameters-selector';
    selector.innerHTML = `
        <h3>🔧 Выберите параметры для проверки</h3>
        
        <div style="margin-bottom: 15px;">
            <button class="select-all-btn" onclick="selectAllParameters()">Выбрать все</button>
            <button class="select-all-btn" onclick="deselectAllParameters()" style="background: #6c757d;">Сбросить все</button>
        </div>
        
        <div class="parameters-category">
            <h4>📐 Форматирование</h4>
            <div class="parameters-list" id="formatting-params"></div>
        </div>
        
        <div class="parameters-category">
            <h4>📏 Интервалы</h4>
            <div class="parameters-list" id="spacing-params"></div>
        </div>
        
        <div class="parameters-category">
            <h4>🔤 Текст</h4>
            <div class="parameters-list" id="text-params"></div>
        </div>
        
        <div class="parameters-category">
            <h4>✒️ Шрифты</h4>
            <div class="parameters-list" id="font-params"></div>
        </div>
        
        <div class="parameters-category">
            <h4>🏷️ Классы</h4>
            <div class="parameters-list" id="class-params"></div>
        </div>
    `;
    
    uploadArea.parentNode.insertBefore(selector, uploadArea.nextSibling);
    
    // Заполняем списки параметров
    allParameters.forEach(param => {
        const container = document.getElementById(`${param.category}-params`);
        if (container) {
            const div = document.createElement('div');
            div.className = 'parameter-checkbox';
            div.innerHTML = `
                <input type="checkbox" id="param-${param.id}" value="${param.id}" checked onchange="toggleParameter('${param.id}')">
                <label for="param-${param.id}">${param.name}</label>
            `;
            container.appendChild(div);
            selectedParameters.add(param.id);
        }
    });
}

function toggleParameter(paramId) {
    if (selectedParameters.has(paramId)) {
        selectedParameters.delete(paramId);
    } else {
        selectedParameters.add(paramId);
    }
}

function selectAllParameters() {
    document.querySelectorAll('.parameter-checkbox input[type="checkbox"]').forEach(cb => {
        cb.checked = true;
        selectedParameters.add(cb.value);
    });
}

function deselectAllParameters() {
    document.querySelectorAll('.parameter-checkbox input[type="checkbox"]').forEach(cb => {
        cb.checked = false;
        selectedParameters.delete(cb.value);
    });
}

async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    // Проверка расширения файла
    if (!file.name.toLowerCase().endsWith('.docx')) {
        showError('Пожалуйста, загрузите файл в формате DOCX');
        return;
    }
    
    updateFileInfo(file);
    showLoading();
    
    try {
        const arrayBuffer = await file.arrayBuffer();
        
        // Используем JSZip для распаковки DOCX
        const zip = await JSZip.loadAsync(arrayBuffer);
        
        // Получаем основной документ
        const documentXml = await zip.file('word/document.xml').async('string');
        
        // Парсим XML
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(documentXml, 'text/xml');
        
        // Анализируем документ по параграфам
        const results = await analyzeDocumentByParagraphs(xmlDoc);
        currentResults = results;
        displayResultsByParagraphs(results);
        
    } catch (error) {
        console.error('Детали ошибки:', error);
        showError(`Ошибка при анализе документа: ${error.message}. Проверьте консоль для деталей.`);
    } finally {
        hideLoading();
    }
}

function updateFileInfo(file) {
    const fileSize = (file.size / 1024).toFixed(2);
    const fileDate = new Date(file.lastModified).toLocaleDateString('ru-RU');
    document.getElementById('fileInfo').innerHTML = `
        <strong>Файл:</strong> ${file.name}<br>
        <strong>Размер:</strong> ${fileSize} KB<br>
        <strong>Изменен:</strong> ${fileDate}
    `;
}

function showLoading() {
    document.getElementById('loading').style.display = 'block';
    document.getElementById('results').style.display = 'none';
}

function hideLoading() {
    document.getElementById('loading').style.display = 'none';
}

function showError(message) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = `<div class="error-message">❌ ${message}</div>`;
    resultsDiv.style.display = 'block';
    document.getElementById('loading').style.display = 'none';
}

async function analyzeDocumentByParagraphs(xmlDoc) {
    // Получаем все параграфы
    const paragraphs = xmlDoc.getElementsByTagName('w:p');
    
    console.log(`Найдено параграфов: ${paragraphs.length}`);
    
    const results = {
        metadata: {
            totalParagraphs: paragraphs.length,
            analysisDate: new Date().toLocaleString('ru-RU'),
            selectedParameters: Array.from(selectedParameters)
        },
        paragraphs: []
    };
    
    // Сначала собираем все параграфы для анализа классов
    const paragraphsData = [];
    Array.from(paragraphs).forEach((paragraph, index) => {
        const text = getParagraphText(paragraph);
        paragraphsData.push({
            element: paragraph,
            text: text,
            hasText: text.trim().length > 0
        });
    });
    
    // Определяем классы для всех параграфов
    const paragraphClasses = determineParagraphClasses(paragraphsData);
    
    // Анализируем каждый параграф
    paragraphsData.forEach((pData, index) => {
        if (pData.hasText) {
            const paragraphData = analyzeParagraph(pData.element, index + 1, paragraphClasses[index]);
            results.paragraphs.push(paragraphData);
        }
    });
    
    // Добавляем сводную статистику
    results.summary = generateSummary(results.paragraphs);
    
    return results;
}

// Новая функция для определения классов параграфов
function determineParagraphClasses(paragraphsData) {
    const classes = [];
    
    for (let i = 0; i < paragraphsData.length; i++) {
        const current = paragraphsData[i];
        const prev = i > 0 ? paragraphsData[i - 1] : null;
        const next = i < paragraphsData.length - 1 ? paragraphsData[i + 1] : null;
        
        let paragraphClass = determineSingleParagraphClass(current, prev, next);
        classes.push(paragraphClass);
    }
    
    return classes;
}

function determineSingleParagraphClass(current, prev, next) {
    const text = current.text.trim();
    
    // Проверка на изображение (g1)
    if (isImage(current.element)) {
        return 'g1';
    }
    
    // Проверка на подпись изображения (h1)
    if (isImageCaption(current, prev)) {
        return 'h1';
    }
    
    // Проверка на подпись таблицы (f1)
    if (isTableCaption(current, prev)) {
        return 'f1';
    }
    
    // Проверка на заголовки
    if (isHeading(text)) {
        if (isHeadingLevel1(text)) return 'b1';
        if (isHeadingLevel2(text)) return 'b2';
        if (isHeadingLevel3(text)) return 'b3';
    }
    
    // Проверка на элементы списка
    if (isListItem(text)) {
        if (!prev || !isListItem(prev.text)) {
            return 'd1'; // Первый элемент списка
        } else if (!next || !isListItem(next.text)) {
            return 'd3'; // Последний элемент списка
        } else {
            return 'd2'; // Средний элемент списка
        }
    }
    
    // Проверка на абзац перед списком (c2)
    if (prev && isListItem(prev.text) && !isListItem(text)) {
        return 'c2';
    }
    
    // Обычный абзац (c1)
    return 'c1';
}

function isImage(paragraph) {
    // Проверяем наличие drawing или pict элементов
    const drawings = paragraph.getElementsByTagName('w:drawing');
    const pictures = paragraph.getElementsByTagName('w:pict');
    return drawings.length > 0 || pictures.length > 0;
}

function isImageCaption(current, prev) {
    if (!prev) return false;
    const prevClass = determineSingleParagraphClass(prev, null, null);
    return prevClass === 'g1' && current.text.toLowerCase().includes('рисунок');
}

function isTableCaption(current, prev) {
    if (!prev) return false;
    return current.text.toLowerCase().includes('таблица');
}

function isHeading(text) {
    const headingPatterns = [
        /^глава\s+\d+/i,
        /^раздел\s+\d+/i,
        /^[IVX]+\./,
        /^\d+\.\s+[А-Я]/,
        /^[А-Я][А-Я\s]+$/  // Заголовок заглавными буквами
    ];
    return headingPatterns.some(pattern => pattern.test(text));
}

function isHeadingLevel1(text) {
    return /^глава\s+\d+/i.test(text) || /^раздел\s+\d+/i.test(text);
}

function isHeadingLevel2(text) {
    return /^\d+\.\s+[А-Я]/.test(text);
}

function isHeadingLevel3(text) {
    return /^\d+\.\d+\.\s+[А-Я]/.test(text);
}

function isListItem(text) {
    const listPatterns = [
        /^[•\-\*]\s/,
        /^\d+\.\s/,
        /^[a-zа-я]\)\s/,
        /^[ivx]+\.\s/i
    ];
    return listPatterns.some(pattern => pattern.test(text));
}

function analyzeParagraph(paragraph, paragraphNumber, paragraphClass) {
    const paragraphData = {
        number: paragraphNumber,
        class: paragraphClass,
        hasText: true,
        text: getParagraphText(paragraph),
        properties: {}
    };
    
    // Получаем все runs в параграфе
    const runs = getRunsInParagraph(paragraph);
    
    // Анализируем шрифты в параграфе
    const fontAnalysis = analyzeFontsInParagraph(runs);
    
    // Собираем только выбранные параметры
    const props = {};
    
    if (selectedParameters.has('alignment')) props.alignment = getAlignment(paragraph);
    if (selectedParameters.has('specialIndentation')) props.specialIndentation = getSpecialIndentation(paragraph);
    if (selectedParameters.has('keepWithNext')) props.keepWithNext = getKeepWithNext(paragraph);
    if (selectedParameters.has('outlineLevel')) props.outlineLevel = getOutlineLevel(paragraph);
    if (selectedParameters.has('keepLinesTogether')) props.keepLinesTogether = getKeepLinesTogether(paragraph);
    if (selectedParameters.has('pageBreakBefore')) props.pageBreakBefore = getPageBreakBefore(paragraph);
    
    if (selectedParameters.has('spaceBefore')) props.spaceBefore = getSpaceBefore(paragraph);
    if (selectedParameters.has('spaceAfter')) props.spaceAfter = getSpaceAfter(paragraph);
    if (selectedParameters.has('lineSpacing')) props.lineSpacing = getLineSpacing(paragraph);
    if (selectedParameters.has('leftIndentation')) props.leftIndentation = getLeftIndentation(paragraph);
    if (selectedParameters.has('rightIndentation')) props.rightIndentation = getRightIndentation(paragraph);
    
    if (selectedParameters.has('fullBold')) props.fullBold = checkFullBold(paragraph, runs);
    if (selectedParameters.has('fullItalic')) props.fullItalic = checkFullItalic(paragraph, runs);
    if (selectedParameters.has('isLowercase')) props.isLowercase = checkLowercase(paragraphData.text);
    if (selectedParameters.has('isUppercase')) props.isUppercase = checkUppercase(paragraphData.text);
    if (selectedParameters.has('specialSymbolsCount')) props.specialSymbolsCount = countSpecialSymbolsInText(paragraphData.text);
    
    if (selectedParameters.has('fonts')) props.fonts = fontAnalysis.fonts;
    if (selectedParameters.has('fontSizes')) props.fontSizes = fontAnalysis.fontSizes;
    if (selectedParameters.has('isTimesNewRoman')) props.isTimesNewRoman = fontAnalysis.isTimesNewRoman;
    if (selectedParameters.has('mainFontSize')) props.mainFontSize = fontAnalysis.mainFontSize;
    if (selectedParameters.has('fontSizeConsistency')) props.fontSizeConsistency = fontAnalysis.sizeConsistency;
    
    // Добавляем класс и вычисленную оценку
    if (selectedParameters.has('class')) props.class = paragraphClass;
    if (selectedParameters.has('computed_mark')) {
        props.computed_mark = calculateComputedMark(paragraphClass, paragraphData.text, runs);
    }
    
    paragraphData.properties = props;
    
    return paragraphData;
}

// Новая функция для вычисления оценки
function calculateComputedMark(paragraphClass, text, runs) {
    let mark = 5; // Начинаем с отлично
    
    // Проверки для разных классов
    switch(paragraphClass) {
        case 'b1':
        case 'b2':
        case 'b3':
            // Заголовки должны быть жирными
            if (!checkFullBold(null, runs)) mark -= 1;
            // Заголовки должны быть по центру
            break;
            
        case 'c1':
            // Обычный абзац должен иметь стандартный шрифт
            if (!checkTimesNewRoman(runs)) mark -= 1;
            break;
            
        case 'd1':
        case 'd2':
        case 'd3':
            // Элементы списка должны иметь правильные отступы
            break;
            
        case 'g1':
            // Изображения должны иметь подпись
            mark = 4; // Условно
            break;
    }
    
    // Проверка на Times New Roman
    if (!checkTimesNewRoman(runs)) mark -= 1;
    
    // Проверка на размер шрифта (должен быть 14pt для обычного текста)
    const fontSize = getMainFontSize(runs);
    if (fontSize && fontSize !== '14pt') mark -= 1;
    
    return Math.max(1, mark); // Не меньше 1
}

function checkTimesNewRoman(runs) {
    for (let run of runs) {
        const font = getRunProperty(run, 'rFonts', 'ascii') || '';
        if (!font.toLowerCase().includes('times')) {
            return false;
        }
    }
    return true;
}

function getMainFontSize(runs) {
    const sizes = {};
    for (let run of runs) {
        const size = getRunProperty(run, 'sz', 'val');
        if (size) {
            const fontSize = Math.round(parseInt(size) / 2) + 'pt';
            sizes[fontSize] = (sizes[fontSize] || 0) + 1;
        }
    }
    
    if (Object.keys(sizes).length > 0) {
        return Object.entries(sizes).sort((a, b) => b[1] - a[1])[0][0];
    }
    return null;
}

// Функция для анализа шрифтов в параграфе
function analyzeFontsInParagraph(runs) {
    const fontsInfo = [];
    const sizesInfo = [];
    const issues = [];
    
    let hasNonTimesNewRoman = false;
    let hasMultipleSizes = false;
    let mainFontSize = null;
    let fontSizeCount = {};
    
    runs.forEach(run => {
        const text = run.getElementsByTagName('w:t')[0]?.textContent || '';
        if (text.trim().length === 0) return;
        
        // Получаем информацию о шрифте
        const font = getRunProperty(run, 'rFonts', 'ascii') || 
                     getRunProperty(run, 'rFonts', 'hAnsi') || 
                     getRunProperty(run, 'rFonts', 'cs') || 
                     'не указан';
        
        // Получаем размер шрифта
        let size = getRunProperty(run, 'sz', 'val');
        
        // Конвертируем размер в пункты
        let fontSize = 'не указан';
        if (size) {
            const sizeNum = parseInt(size);
            if (!isNaN(sizeNum) && sizeNum > 0) {
                fontSize = Math.round(sizeNum / 2) + 'pt';
                fontSizeCount[fontSize] = (fontSizeCount[fontSize] || 0) + 1;
            }
        }
        
        fontsInfo.push(font);
        sizesInfo.push(fontSize);
        
        // Проверяем, является ли шрифт Times New Roman
        const isTimesNewRoman = font.toLowerCase().includes('times') || 
                                font === 'Times New Roman' ||
                                font === 'Times';
        
        if (!isTimesNewRoman && font !== 'не указан') {
            hasNonTimesNewRoman = true;
        }
    });
    
    // Определяем основной размер
    if (Object.keys(fontSizeCount).length > 0) {
        mainFontSize = Object.entries(fontSizeCount)
            .sort((a, b) => b[1] - a[1])[0][0];
    }
    
    // Проверяем на разные размеры
    const uniqueSizes = [...new Set(sizesInfo.filter(s => s !== 'не указан'))];
    if (uniqueSizes.length > 1) {
        hasMultipleSizes = true;
        issues.push(`Разные размеры: ${uniqueSizes.join(', ')}`);
    }
    
    const uniqueFonts = [...new Set(fontsInfo)];
    
    return {
        fonts: uniqueFonts.join(', '),
        fontSizes: [...new Set(sizesInfo)].join(', '),
        isTimesNewRoman: !hasNonTimesNewRoman && uniqueFonts.length > 0,
        issues: issues,
        sizeConsistency: !hasMultipleSizes,
        mainFontSize: mainFontSize
    };
}

// Вспомогательные функции для извлечения данных
function getParagraphText(paragraph) {
    let text = '';
    try {
        const texts = paragraph.getElementsByTagName('w:t');
        for (let t of texts) {
            if (t.textContent) {
                text += t.textContent;
            }
        }
    } catch (e) {
        console.warn('Ошибка при извлечении текста параграфа:', e);
    }
    return text;
}

function getRunsInParagraph(paragraph) {
    try {
        return Array.from(paragraph.getElementsByTagName('w:r'));
    } catch (e) {
        return [];
    }
}

function getParagraphProperty(paragraph, propName, subProp = null) {
    try {
        const pPr = paragraph.getElementsByTagName('w:pPr')[0];
        if (!pPr) return null;
        
        const prop = pPr.getElementsByTagName(`w:${propName}`)[0];
        if (!prop) return null;
        
        if (subProp) {
            const attr = prop.getAttribute(`w:${subProp}`);
            if (attr !== null) return attr;
            
            const subElement = prop.getElementsByTagName(`w:${subProp}`)[0];
            if (subElement) {
                return subElement.getAttribute('w:val');
            }
            return null;
        }
        
        return true;
    } catch (e) {
        return null;
    }
}

function getRunProperty(run, propName, subProp = null) {
    try {
        const rPr = run.getElementsByTagName('w:rPr')[0];
        if (!rPr) return null;
        
        const prop = rPr.getElementsByTagName(`w:${propName}`)[0];
        if (!prop) return null;
        
        if (subProp) {
            const attr = prop.getAttribute(`w:${subProp}`);
            if (attr !== null && attr !== '') return attr;
            
            const subElement = prop.getElementsByTagName(`w:${subProp}`)[0];
            if (subElement) {
                return subElement.getAttribute('w:val');
            }
            return null;
        }
        
        return true;
    } catch (e) {
        return null;
    }
}

// Функции для анализа параметров параграфа
function getAlignment(paragraph) {
    const alignment = getParagraphProperty(paragraph, 'jc') || 'left';
    const alignmentMap = {
        'left': 'По левому краю',
        'center': 'По центру',
        'right': 'По правому краю',
        'both': 'По ширине'
    };
    return alignmentMap[alignment] || alignment;
}

function getSpecialIndentation(paragraph) {
    const firstLineIndent = getParagraphProperty(paragraph, 'ind', 'firstLine');
    const hangingIndent = getParagraphProperty(paragraph, 'ind', 'hanging');
    
    if (firstLineIndent || hangingIndent) {
        const indentInfo = [];
        if (firstLineIndent) {
            indentInfo.push(`отступ первой строки: ${Math.round(parseInt(firstLineIndent) / 20)} pt`);
        }
        if (hangingIndent) {
            indentInfo.push(`висячий отступ: ${Math.round(parseInt(hangingIndent) / 20)} pt`);
        }
        return indentInfo.join(', ');
    }
    
    return 'Нет';
}

function getKeepWithNext(paragraph) {
    const keep = getParagraphProperty(paragraph, 'keepNext');
    return keep ? 'Да' : 'Нет';
}

function getOutlineLevel(paragraph) {
    const level = getParagraphProperty(paragraph, 'outlineLvl', 'val');
    return level ? `Уровень ${level}` : 'Не задан';
}

function getSpaceBefore(paragraph) {
    const space = getParagraphProperty(paragraph, 'spacing', 'before');
    if (space && space !== '0') {
        return `${Math.round(parseInt(space) / 20)} pt`;
    }
    return '0 pt';
}

function getSpaceAfter(paragraph) {
    const space = getParagraphProperty(paragraph, 'spacing', 'after');
    if (space && space !== '0') {
        return `${Math.round(parseInt(space) / 20)} pt`;
    }
    return '0 pt';
}

function getKeepLinesTogether(paragraph) {
    const keep = getParagraphProperty(paragraph, 'keepLines');
    return keep ? 'Да' : 'Нет';
}

function getLineSpacing(paragraph) {
    const lineRule = getParagraphProperty(paragraph, 'spacing', 'lineRule');
    const line = getParagraphProperty(paragraph, 'spacing', 'line');
    
    if (line) {
        const lineValue = parseInt(line) / 240;
        return `${lineValue.toFixed(1)} (${lineRule || 'auto'})`;
    }
    
    return 'одинарный';
}

function getLeftIndentation(paragraph) {
    const indent = getParagraphProperty(paragraph, 'ind', 'left');
    if (indent && indent !== '0') {
        return `${Math.round(parseInt(indent) / 20)} pt`;
    }
    return '0 pt';
}

function getRightIndentation(paragraph) {
    const indent = getParagraphProperty(paragraph, 'ind', 'right');
    if (indent && indent !== '0') {
        return `${Math.round(parseInt(indent) / 20)} pt`;
    }
    return '0 pt';
}

function getPageBreakBefore(paragraph) {
    const pageBreak = getParagraphProperty(paragraph, 'pageBreakBefore');
    return pageBreak ? 'Да' : 'Нет';
}

function checkFullBold(paragraph, runs) {
    if (runs.length === 0) return 'Нет текста';
    
    let allBold = true;
    let hasText = false;
    
    runs.forEach(run => {
        const text = run.getElementsByTagName('w:t')[0]?.textContent || '';
        if (text.trim().length > 0) {
            hasText = true;
            const bold = getRunProperty(run, 'b');
            if (!bold) allBold = false;
        }
    });
    
    if (!hasText) return 'Нет текста';
    return allBold ? 'Да' : 'Нет';
}

function checkFullItalic(paragraph, runs) {
    if (runs.length === 0) return 'Нет текста';
    
    let allItalic = true;
    let hasText = false;
    
    runs.forEach(run => {
        const text = run.getElementsByTagName('w:t')[0]?.textContent || '';
        if (text.trim().length > 0) {
            hasText = true;
            const italic = getRunProperty(run, 'i');
            if (!italic) allItalic = false;
        }
    });
    
    if (!hasText) return 'Нет текста';
    return allItalic ? 'Да' : 'Нет';
}

function checkLowercase(text) {
    if (text.length === 0) return 'Нет текста';
    if (text === text.toLowerCase() && /[a-zа-я]/.test(text)) {
        return 'Да';
    }
    return 'Нет';
}

function checkUppercase(text) {
    if (text.length === 0) return 'Нет текста';
    if (text === text.toUpperCase() && /[A-ZА-Я]/.test(text)) {
        return 'Да';
    }
    return 'Нет';
}

function countSpecialSymbolsInText(text) {
    const specialSymbols = /[!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?]/g;
    const matches = text.match(specialSymbols);
    return matches ? matches.length : 0;
}

function generateSummary(paragraphs) {
    const summary = {
        totalParagraphsWithText: paragraphs.length,
        classStats: {},
        alignmentStats: {},
        timesNewRomanStats: {
            paragraphsWithTimesNewRoman: 0,
            paragraphsWithIssues: 0
        },
        averageMark: 0,
        classDistribution: {}
    };
    
    let totalMark = 0;
    
    paragraphs.forEach(p => {
        const props = p.properties;
        
        // Статистика по классам
        if (p.class) {
            summary.classStats[p.class] = (summary.classStats[p.class] || 0) + 1;
        }
        
        // Статистика по выравниванию
        if (props.alignment) {
            summary.alignmentStats[props.alignment] = (summary.alignmentStats[props.alignment] || 0) + 1;
        }
        
        // Статистика по Times New Roman
        if (props.isTimesNewRoman !== undefined) {
            if (props.isTimesNewRoman) {
                summary.timesNewRomanStats.paragraphsWithTimesNewRoman++;
            } else {
                summary.timesNewRomanStats.paragraphsWithIssues++;
            }
        }
        
        // Сумма оценок
        if (props.computed_mark) {
            totalMark += props.computed_mark;
        }
    });
    
    summary.averageMark = paragraphs.length > 0 ? (totalMark / paragraphs.length).toFixed(2) : 0;
    
    return summary;
}

function displayResultsByParagraphs(results) {
    document.getElementById('results').style.display = 'block';
    
    // Отображаем сводку
    const summary = document.getElementById('summary');
    summary.innerHTML = `
        <h3>📊 Сводка анализа</h3>
        <p>📄 Всего параграфов с текстом: <strong>${results.summary.totalParagraphsWithText}</strong></p>
        <p>📅 Дата анализа: <strong>${results.metadata.analysisDate}</strong></p>
        <p>✅ Статус: <span class="success">Анализ завершен успешно</span></p>
        <p>📊 Средняя оценка: <strong>${results.summary.averageMark}</strong></p>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px; margin-top: 20px;">
            <div style="background: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h4 style="color: #667eea; margin-bottom: 10px;">🏷️ Распределение классов</h4>
                ${Object.entries(results.summary.classStats).map(([className, count]) => 
                    `<p><span class="class-badge ${className}">${className}</span> <strong>${count}</strong></p>`
                ).join('')}
            </div>
            
            <div style="background: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                <h4 style="color: #667eea; margin-bottom: 10px;">🔤 Times New Roman</h4>
                <p><strong>Соответствуют:</strong> ${results.summary.timesNewRomanStats.paragraphsWithTimesNewRoman}</p>
                <p><strong>Не соответствуют:</strong> ${results.summary.timesNewRomanStats.paragraphsWithIssues}</p>
            </div>
        </div>
    `;
    
    // Отображаем параграфы
    const grid = document.getElementById('parametersGrid');
    grid.innerHTML = '<h3 style="grid-column: 1/-1; margin: 20px 0 10px;">📑 Детальный анализ по параграфам</h3>';
    
    results.paragraphs.forEach(paragraph => {
        const card = createParagraphCard(paragraph);
        grid.appendChild(card);
    });
    
    // Добавляем кнопку экспорта
    const exportBtn = document.querySelector('.export-btn');
    exportBtn.style.display = 'block';
}

function createParagraphCard(paragraph) {
    const card = document.createElement('div');
    card.className = 'parameter-card';
    card.style.gridColumn = 'span 1';
    
    // Обрезаем текст для отображения
    const displayText = paragraph.text.length > 150 
        ? paragraph.text.substring(0, 150) + '...' 
        : paragraph.text;
    
    const props = paragraph.properties;
    
    // Собираем HTML для свойств на основе выбранных параметров
    let propertiesHtml = '';
    
    // Группируем свойства по категориям
    if (props.alignment || props.specialIndentation || props.keepWithNext || props.outlineLevel || 
        props.keepLinesTogether || props.pageBreakBefore) {
        propertiesHtml += '<div><h4 style="color: #667eea; font-size: 0.9em; margin-bottom: 5px;">📐 Форматирование</h4><ul style="list-style: none; padding: 0; font-size: 0.85em;">';
        if (props.alignment) propertiesHtml += `<li><strong>Выравнивание:</strong> ${props.alignment}</li>`;
        if (props.specialIndentation) propertiesHtml += `<li><strong>Отступы:</strong> ${props.specialIndentation}</li>`;
        if (props.keepWithNext) propertiesHtml += `<li><strong>Связать со след.:</strong> ${props.keepWithNext}</li>`;
        if (props.outlineLevel) propertiesHtml += `<li><strong>Уровень:</strong> ${props.outlineLevel}</li>`;
        if (props.keepLinesTogether) propertiesHtml += `<li><strong>Не разрывать:</strong> ${props.keepLinesTogether}</li>`;
        if (props.pageBreakBefore) propertiesHtml += `<li><strong>Разрыв страницы:</strong> ${props.pageBreakBefore}</li>`;
        propertiesHtml += '</ul></div>';
    }
    
    if (props.spaceBefore || props.spaceAfter || props.lineSpacing || props.leftIndentation || props.rightIndentation) {
        propertiesHtml += '<div><h4 style="color: #667eea; font-size: 0.9em; margin-bottom: 5px;">📏 Интервалы</h4><ul style="list-style: none; padding: 0; font-size: 0.85em;">';
        if (props.spaceBefore) propertiesHtml += `<li><strong>Перед:</strong> ${props.spaceBefore}</li>`;
        if (props.spaceAfter) propertiesHtml += `<li><strong>После:</strong> ${props.spaceAfter}</li>`;
        if (props.lineSpacing) propertiesHtml += `<li><strong>Межстрочный:</strong> ${props.lineSpacing}</li>`;
        if (props.leftIndentation) propertiesHtml += `<li><strong>Левый отступ:</strong> ${props.leftIndentation}</li>`;
        if (props.rightIndentation) propertiesHtml += `<li><strong>Правый отступ:</strong> ${props.rightIndentation}</li>`;
        propertiesHtml += '</ul></div>';
    }
    
    if (props.fullBold || props.fullItalic || props.isLowercase || props.isUppercase || props.specialSymbolsCount !== undefined) {
        propertiesHtml += '<div><h4 style="color: #667eea; font-size: 0.9em; margin-bottom: 5px;">🔤 Текст</h4><ul style="list-style: none; padding: 0; font-size: 0.85em;">';
        if (props.fullBold) propertiesHtml += `<li><strong>Жирный:</strong> ${props.fullBold}</li>`;
        if (props.fullItalic) propertiesHtml += `<li><strong>Курсив:</strong> ${props.fullItalic}</li>`;
        if (props.isLowercase) propertiesHtml += `<li><strong>Строчные:</strong> ${props.isLowercase}</li>`;
        if (props.isUppercase) propertiesHtml += `<li><strong>Прописные:</strong> ${props.isUppercase}</li>`;
        if (props.specialSymbolsCount !== undefined) propertiesHtml += `<li><strong>Спецсимволы:</strong> ${props.specialSymbolsCount}</li>`;
        propertiesHtml += '</ul></div>';
    }
    
    if (props.fonts || props.fontSizes || props.isTimesNewRoman !== undefined || props.mainFontSize || props.fontSizeConsistency !== undefined) {
        propertiesHtml += '<div><h4 style="color: #667eea; font-size: 0.9em; margin-bottom: 5px;">✒️ Шрифты</h4><ul style="list-style: none; padding: 0; font-size: 0.85em;">';
        if (props.fonts) propertiesHtml += `<li><strong>Шрифты:</strong> ${props.fonts}</li>`;
        if (props.fontSizes) propertiesHtml += `<li><strong>Размеры:</strong> ${props.fontSizes}</li>`;
        if (props.isTimesNewRoman !== undefined) propertiesHtml += `<li><strong>Times New Roman:</strong> ${props.isTimesNewRoman ? '✅' : '❌'}</li>`;
        if (props.mainFontSize) propertiesHtml += `<li><strong>Осн. размер:</strong> ${props.mainFontSize}</li>`;
        if (props.fontSizeConsistency !== undefined) propertiesHtml += `<li><strong>Единый размер:</strong> ${props.fontSizeConsistency ? '✅' : '❌'}</li>`;
        propertiesHtml += '</ul></div>';
    }
    
    card.innerHTML = `
        <div class="parameter-name" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
            <span>📄 Параграф ${paragraph.number}</span>
            <div>
                <span class="class-badge ${paragraph.class}">${paragraph.class}</span>
                ${props.computed_mark ? `<span style="margin-left: 10px; font-size: 0.9em; background: ${getMarkColor(props.computed_mark)}; color: white; padding: 3px 8px; border-radius: 12px;">Оценка: ${props.computed_mark}</span>` : ''}
            </div>
        </div>
        
        <div style="margin-bottom: 15px; padding: 10px; background: #f8f9fa; border-radius: 8px; font-style: italic; color: #666; max-height: 100px; overflow-y: auto;">
            "${displayText}"
        </div>
        
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
            ${propertiesHtml}
        </div>
    `;
    
    return card;
}

function getMarkColor(mark) {
    if (mark >= 4.5) return '#27ae60';
    if (mark >= 3.5) return '#f39c12';
    return '#e74c3c';
}

// Функция для экспорта результатов
window.exportResults = function() {
    if (!currentResults) {
        alert('Нет данных для экспорта');
        return;
    }
    
    const dataStr = JSON.stringify(currentResults, null, 2);
    const dataUri = 'data:application/json;charset=utf-8,'+ encodeURIComponent(dataStr);
    
    const date = new Date();
    const fileName = `docx_analysis_${date.getFullYear()}-${(date.getMonth()+1).toString().padStart(2,'0')}-${date.getDate().toString().padStart(2,'0')}.json`;
    
    const linkElement = document.createElement('a');
    linkElement.setAttribute('href', dataUri);
    linkElement.setAttribute('download', fileName);
    linkElement.click();
};

// Делаем функции глобальными
window.toggleParameter = toggleParameter;
window.selectAllParameters = selectAllParameters;
window.deselectAllParameters = deselectAllParameters;