// 문서 글자수 카운터 앱
class DocumentTextCounter {
    constructor() {
        this.dropZone = document.getElementById('dropZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileTypeIcon = document.getElementById('fileTypeIcon');
        this.progressSection = document.getElementById('progressSection');
        this.progressFill = document.getElementById('progressFill');
        this.progressText = document.getElementById('progressText');
        this.resultsSection = document.getElementById('resultsSection');
        this.errorMessage = document.getElementById('errorMessage');
        this.slidesContainer = document.getElementById('slidesContainer');
        this.sectionTitle = document.getElementById('sectionTitle');
        this.languageGrid = document.getElementById('languageGrid');

        this.imageFormats = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp'];
        this.supportedFormats = ['pptx', 'docx', 'xlsx', 'pdf', 'txt', ...this.imageFormats];
        this.oldFormats = ['ppt', 'doc', 'xls'];

        // 기본 언어: 한국어, 영어
        this.defaultLanguages = ['kor', 'eng'];

        this.initEventListeners();
        this.loadLanguageSettings();
    }

    initEventListeners() {
        this.dropZone.addEventListener('click', () => this.fileInput.click());

        this.fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                this.handleFile(e.target.files[0]);
            }
        });

        this.dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            this.dropZone.classList.add('dragover');
        });

        this.dropZone.addEventListener('dragleave', () => {
            this.dropZone.classList.remove('dragover');
        });

        this.dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            this.dropZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                this.handleFile(e.dataTransfer.files[0]);
            }
        });

        // 언어 선택 변경 시 저장
        this.languageGrid.addEventListener('change', () => {
            this.saveLanguageSettings();
        });
    }

    // 언어 설정 로드
    loadLanguageSettings() {
        const saved = localStorage.getItem('ocrLanguages');
        const languages = saved ? JSON.parse(saved) : this.defaultLanguages;

        // 모든 체크박스 초기화
        const checkboxes = this.languageGrid.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(checkbox => {
            checkbox.checked = languages.includes(checkbox.value);
        });
    }

    // 언어 설정 저장
    saveLanguageSettings() {
        const languages = this.getSelectedLanguages();
        localStorage.setItem('ocrLanguages', JSON.stringify(languages));
    }

    // 선택된 언어 가져오기
    getSelectedLanguages() {
        const checkboxes = this.languageGrid.querySelectorAll('input[type="checkbox"]:checked');
        const languages = Array.from(checkboxes).map(cb => cb.value);
        // 최소 1개 이상 선택되어야 함
        return languages.length > 0 ? languages : this.defaultLanguages;
    }

    async handleFile(file) {
        const ext = file.name.toLowerCase().split('.').pop();

        if (this.oldFormats.includes(ext)) {
            this.showError(`구 버전 ${ext.toUpperCase()} 파일은 브라우저에서 직접 처리할 수 없습니다. ${ext.toUpperCase()}X 형식으로 변환 후 다시 시도해주세요.`);
            return;
        }

        if (!this.supportedFormats.includes(ext)) {
            this.showError('지원하지 않는 파일 형식입니다. PPTX, DOCX, XLSX, PDF 또는 이미지 파일을 선택해주세요.');
            return;
        }

        // UI 초기화
        this.hideError();
        this.fileInfo.classList.add('show');
        this.fileName.textContent = file.name;

        // 파일 타입 아이콘 표시
        const iconType = this.imageFormats.includes(ext) ? 'img' : ext;
        this.fileTypeIcon.textContent = ext.toUpperCase();
        this.fileTypeIcon.className = `file-type-icon file-type-${iconType}`;

        this.progressSection.classList.add('show');
        this.resultsSection.classList.remove('show');
        this.updateProgress(0, '파일 읽는 중...');

        try {
            let results;
            if (ext === 'pptx') {
                results = await this.analyzePPTX(file);
                this.sectionTitle.textContent = '슬라이드별 상세 정보';
            } else if (ext === 'docx') {
                results = await this.analyzeDOCX(file);
                this.sectionTitle.textContent = '문서 상세 정보';
            } else if (ext === 'xlsx') {
                results = await this.analyzeXLSX(file);
                this.sectionTitle.textContent = '시트별 상세 정보';
            } else if (ext === 'pdf') {
                results = await this.analyzePDF(file);
                this.sectionTitle.textContent = '페이지별 상세 정보';
            } else if (ext === 'txt') {
                results = await this.analyzeTXT(file);
                this.sectionTitle.textContent = '텍스트 파일 정보';
            } else if (this.imageFormats.includes(ext)) {
                results = await this.analyzeImage(file);
                this.sectionTitle.textContent = '이미지 분석 결과';
            }
            this.displayResults(results);
        } catch (error) {
            console.error('분석 오류:', error);
            this.showError(`파일 분석 중 오류가 발생했습니다: ${error.message}`);
            this.progressSection.classList.remove('show');
        }
    }

    // ============ 이미지 전처리 함수 ============
    // 아랍어, 힌디어 등 연결 문자 체계를 위해 최소한의 전처리만 수행
    async preprocessImageForOCR(imageData) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');

                // 이미지 크기 설정
                canvas.width = img.width;
                canvas.height = img.height;

                // 흰색 배경 설정 (투명 배경 문제 방지)
                ctx.fillStyle = '#FFFFFF';
                ctx.fillRect(0, 0, canvas.width, canvas.height);

                // 원본 이미지 그리기
                ctx.drawImage(img, 0, 0);

                // 이미지 데이터 가져오기
                const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                const data = imgData.data;

                // 가벼운 대비 향상만 적용 (이진화 제거 - 연결 문자 보존)
                const contrastFactor = 1.2;
                for (let i = 0; i < data.length; i += 4) {
                    // 대비 향상
                    for (let c = 0; c < 3; c++) {
                        let val = ((data[i + c] - 128) * contrastFactor) + 128;
                        data[i + c] = Math.max(0, Math.min(255, val));
                    }
                }

                ctx.putImageData(imgData, 0, 0);
                resolve(canvas.toDataURL('image/png'));
            };
            img.onerror = () => resolve(imageData); // 실패시 원본 반환
            img.src = imageData;
        });
    }

    async scaleImageForOCR(imageData, minSize = 300) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const { width, height } = img;

                // 이미 충분히 크면 그대로 반환
                if (width >= minSize && height >= minSize) {
                    resolve(imageData);
                    return;
                }

                // 스케일 계산 (최소 크기 보장)
                const scale = Math.max(minSize / width, minSize / height, 1);
                const newWidth = Math.round(width * scale);
                const newHeight = Math.round(height * scale);

                const canvas = document.createElement('canvas');
                canvas.width = newWidth;
                canvas.height = newHeight;
                const ctx = canvas.getContext('2d');

                // 고품질 스케일링
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(img, 0, 0, newWidth, newHeight);

                resolve(canvas.toDataURL('image/png'));
            };
            img.onerror = () => resolve(imageData);
            img.src = imageData;
        });
    }

    // OCR 입력 이미지 크기 정규화 (모든 포맷에서 동일한 결과를 위해)
    async normalizeImageForOCR(imageData) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                const { width, height } = img;

                // 목표: 장축이 3000px이 되도록 스케일링 (OCR 정확도 향상)
                const targetMaxDim = 3000;
                const maxDim = Math.max(width, height);
                const scale = targetMaxDim / maxDim;

                const newWidth = Math.round(width * scale);
                const newHeight = Math.round(height * scale);

                const canvas = document.createElement('canvas');
                canvas.width = newWidth;
                canvas.height = newHeight;
                const ctx = canvas.getContext('2d');

                // 고품질 스케일링
                ctx.imageSmoothingEnabled = true;
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(img, 0, 0, newWidth, newHeight);

                resolve(canvas.toDataURL('image/png'));
            };
            img.onerror = () => resolve(imageData);
            img.src = imageData;
        });
    }

    getOptimalPDFScale(page) {
        const viewport = page.getViewport({ scale: 1 });
        const { width, height } = viewport;

        // 목표: 300 DPI 품질 (A4 기준 약 2480x3508 픽셀)
        // 최소 스케일 2, 최대 스케일 4
        const targetPixels = 2000;
        const currentMax = Math.max(width, height);

        let scale = targetPixels / currentMax;
        scale = Math.max(2, Math.min(4, scale)); // 2~4 범위로 제한

        return scale;
    }

    // ============ 공용 OCR 처리 함수 ============
    async processImagesOCR(images, progressCallback) {
        let ocrText = '';
        let ocrCharCount = 0;

        for (let i = 0; i < images.length; i++) {
            if (progressCallback) {
                progressCallback(i, images.length);
            }

            try {
                const imageData = typeof images[i] === 'string' ? images[i] : images[i].data;
                const result = await this.performOCR(imageData);
                if (result) {
                    ocrText += result + '\n';
                }
            } catch (err) {
                console.warn('OCR 오류:', err);
            }
        }

        ocrCharCount = this.countChars(ocrText);
        return { ocrText: ocrText.trim(), ocrCharCount };
    }

    async performOCR(imageData) {
        try {
            // 1. 이미지 크기 정규화 (모든 포맷에서 동일한 크기로)
            const normalizedImage = await this.normalizeImageForOCR(imageData);

            // 2. 이미지 전처리 (대비 향상)
            const processedImage = await this.preprocessImageForOCR(normalizedImage);

            // 3. Tesseract OCR 실행 (선택된 언어만 사용)
            const selectedLanguages = this.getSelectedLanguages();
            const langString = selectedLanguages.join('+');

            const result = await Tesseract.recognize(
                processedImage,
                langString,
                {
                    logger: () => {},
                }
            );
            return result.data.text || '';
        } catch (error) {
            console.warn('OCR 처리 실패:', error);
            return '';
        }
    }

    // ============ TXT 파일 분석 ============
    async analyzeTXT(file) {
        this.updateProgress(10, '텍스트 파일 읽는 중...');

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        const text = await file.text();
        const textCharCount = this.countChars(text);

        results.items.push({
            name: '본문',
            textCharCount,
            ocrText: '',
            ocrCharCount: 0,
            totalCharCount: textCharCount,
            imageCount: 0
        });

        results.totalTextChars = textCharCount;
        results.totalChars = textCharCount;

        this.updateProgress(100, '분석 완료!');
        return results;
    }

    // ============ 이미지 파일 분석 ============
    async analyzeImage(file) {
        this.updateProgress(10, '이미지 로딩 중...');

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        // 파일을 Data URL로 변환
        const imageData = await this.fileToDataURL(file);

        this.updateProgress(30, 'OCR 처리 중...');

        const { ocrText, ocrCharCount } = await this.processImagesOCR([imageData], (i, total) => {
            this.updateProgress(30 + (i / total) * 60, `OCR 처리 중...`);
        });

        results.items.push({
            name: '이미지',
            textCharCount: 0,
            ocrText: ocrText,
            ocrCharCount: ocrCharCount,
            totalCharCount: ocrCharCount,
            imageCount: 1
        });

        results.totalOcrChars = ocrCharCount;
        results.totalChars = ocrCharCount;

        this.updateProgress(100, '분석 완료!');
        return results;
    }

    fileToDataURL(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    // ============ PPTX 분석 ============
    async analyzePPTX(file) {
        const zip = new JSZip();
        const content = await zip.loadAsync(file);

        this.updateProgress(10, '슬라이드 정보 추출 중...');

        const slideFiles = this.getSortedFiles(content, /ppt\/slides\/slide(\d+)\.xml$/);

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        const images = await this.extractImagesFromZip(content, 'ppt/media/');
        this.updateProgress(20, `${images.length}개의 이미지 발견...`);

        // 이미지 데이터 해시로 중복 제거 (같은 내용의 이미지는 한 번만 처리)
        const ocrCache = new Map(); // 해시 -> OCR 결과
        const imageHashMap = new Map(); // 이미지명 -> 해시

        // 모든 이미지의 해시 계산 (픽셀 기반 - 크기 다른 동일 이미지도 감지)
        for (const img of images) {
            const hash = await this.hashImageData(img.data);
            imageHashMap.set(img.name, hash);
        }

        const slideImageMap = await this.mapImagesToSlides(content, slideFiles, images);

        const totalSteps = slideFiles.length;
        for (let i = 0; i < slideFiles.length; i++) {
            const slideFile = slideFiles[i];
            const slideNum = i + 1;

            this.updateProgress(
                20 + (i / totalSteps) * 70,
                `슬라이드 ${slideNum}/${totalSteps} 분석 중...`
            );

            const xmlContent = await content.files[slideFile].async('text');
            const textContent = this.extractTextFromXML(xmlContent, 'a:t');
            const textCharCount = this.countChars(textContent);

            const slideImages = slideImageMap[slideFile] || [];
            let ocrText = '';
            let ocrCharCount = 0;

            if (slideImages.length > 0) {
                // 이 슬라이드의 고유 이미지만 처리 (해시 기반 중복 제거)
                const processedHashes = new Set();
                const ocrResults = [];

                for (let j = 0; j < slideImages.length; j++) {
                    const img = slideImages[j];
                    const imgHash = imageHashMap.get(img.name);

                    // 이미 이 슬라이드에서 처리한 동일 이미지 스킵
                    if (processedHashes.has(imgHash)) {
                        continue;
                    }
                    processedHashes.add(imgHash);

                    if (ocrCache.has(imgHash)) {
                        ocrResults.push(ocrCache.get(imgHash));
                    } else {
                        this.updateProgress(
                            20 + (i / totalSteps) * 70,
                            `슬라이드 ${slideNum} - 이미지 OCR...`
                        );
                        const imageData = typeof img === 'string' ? img : img.data;
                        const result = await this.performOCR(imageData);
                        ocrCache.set(imgHash, result);
                        ocrResults.push(result);
                    }
                }
                ocrText = ocrResults.filter(t => t).join('\n').trim();
                ocrCharCount = this.countChars(ocrText);
            }

            results.items.push({
                name: `슬라이드 ${slideNum}`,
                textCharCount,
                ocrText,
                ocrCharCount,
                totalCharCount: textCharCount + ocrCharCount,
                imageCount: slideImages.length
            });

            results.totalTextChars += textCharCount;
            results.totalOcrChars += ocrCharCount;
        }

        results.totalChars = results.totalTextChars + results.totalOcrChars;
        this.updateProgress(100, '분석 완료!');
        return results;
    }

    // 이미지 데이터의 픽셀 기반 해시 생성 (중복 감지용)
    // 이미지를 32x32로 축소 후 픽셀 데이터로 해시 생성
    // → 크기가 다른 동일 이미지도 같은 해시값
    async hashImageData(dataUrl) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => {
                // 고정 크기로 축소
                const canvas = document.createElement('canvas');
                canvas.width = 32;
                canvas.height = 32;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0, 32, 32);

                // 픽셀 데이터로 해시 생성
                const imageData = ctx.getImageData(0, 0, 32, 32);
                let hash = '';
                // 256개 샘플 포인트에서 픽셀 값 추출
                for (let i = 0; i < imageData.data.length; i += 16) {
                    hash += imageData.data[i].toString(16).padStart(2, '0');
                }
                resolve(hash);
            };
            img.onerror = () => {
                // 실패 시 base64 길이로 폴백
                const base64 = dataUrl.split(',')[1] || dataUrl;
                resolve('err_' + base64.length.toString());
            };
            img.src = dataUrl;
        });
    }

    async mapImagesToSlides(zip, slideFiles, images) {
        const mapping = {};

        for (const slideFile of slideFiles) {
            mapping[slideFile] = [];
            const addedImageNames = new Set(); // 슬라이드별 중복 방지

            const slideNum = slideFile.match(/slide(\d+)\.xml$/)[1];
            const relsFile = `ppt/slides/_rels/slide${slideNum}.xml.rels`;

            if (zip.files[relsFile]) {
                const relsContent = await zip.files[relsFile].async('text');
                const parser = new DOMParser();
                const relsDoc = parser.parseFromString(relsContent, 'text/xml');
                const relationships = relsDoc.getElementsByTagName('Relationship');

                for (const rel of relationships) {
                    const target = rel.getAttribute('Target');
                    if (target && target.includes('../media/')) {
                        const imageName = target.replace('../media/', '');
                        // 이미 추가된 이미지는 스킵
                        if (addedImageNames.has(imageName)) {
                            continue;
                        }
                        const image = images.find(img => img.name === imageName);
                        if (image) {
                            mapping[slideFile].push(image);
                            addedImageNames.add(imageName);
                        }
                    }
                }
            }
        }

        return mapping;
    }

    // ============ DOCX 분석 ============
    async analyzeDOCX(file) {
        const zip = new JSZip();
        const content = await zip.loadAsync(file);

        this.updateProgress(10, '문서 내용 추출 중...');

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        if (content.files['word/document.xml']) {
            const xmlContent = await content.files['word/document.xml'].async('text');
            const textContent = this.extractTextFromXML(xmlContent, 'w:t');
            const textCharCount = this.countChars(textContent);

            results.items.push({
                name: '본문',
                textCharCount,
                ocrText: '',
                ocrCharCount: 0,
                totalCharCount: textCharCount,
                imageCount: 0
            });

            results.totalTextChars += textCharCount;
        }

        this.updateProgress(30, '이미지 추출 중...');

        const images = await this.extractImagesFromZip(content, 'word/media/');

        if (images.length > 0) {
            const { ocrText, ocrCharCount } = await this.processImagesOCR(images, (i, total) => {
                this.updateProgress(30 + (i / total) * 60, `이미지 OCR (${i + 1}/${total})...`);
            });

            if (ocrCharCount > 0) {
                results.items.push({
                    name: '이미지 내 텍스트',
                    textCharCount: 0,
                    ocrText,
                    ocrCharCount,
                    totalCharCount: ocrCharCount,
                    imageCount: images.length
                });

                results.totalOcrChars += ocrCharCount;
            }
        }

        results.totalChars = results.totalTextChars + results.totalOcrChars;
        this.updateProgress(100, '분석 완료!');
        return results;
    }

    // ============ XLSX 분석 ============
    async analyzeXLSX(file) {
        const zip = new JSZip();
        const content = await zip.loadAsync(file);

        this.updateProgress(10, '시트 정보 추출 중...');

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        let sharedStrings = [];
        if (content.files['xl/sharedStrings.xml']) {
            const ssXml = await content.files['xl/sharedStrings.xml'].async('text');
            sharedStrings = this.parseSharedStrings(ssXml);
        }

        const sheetNames = await this.getSheetNames(content);
        const sheetFiles = this.getSortedFiles(content, /xl\/worksheets\/sheet(\d+)\.xml$/);

        const totalSteps = sheetFiles.length;
        for (let i = 0; i < sheetFiles.length; i++) {
            const sheetFile = sheetFiles[i];
            const sheetName = sheetNames[i] || `시트 ${i + 1}`;

            this.updateProgress(10 + (i / totalSteps) * 70, `${sheetName} 분석 중...`);

            const xmlContent = await content.files[sheetFile].async('text');
            const textContent = this.extractXLSXText(xmlContent, sharedStrings);
            const textCharCount = this.countChars(textContent);

            results.items.push({
                name: sheetName,
                textCharCount,
                ocrText: '',
                ocrCharCount: 0,
                totalCharCount: textCharCount,
                imageCount: 0
            });

            results.totalTextChars += textCharCount;
        }

        this.updateProgress(80, '이미지 추출 중...');
        const images = await this.extractImagesFromZip(content, 'xl/media/');

        if (images.length > 0) {
            const { ocrText, ocrCharCount } = await this.processImagesOCR(images, (i, total) => {
                this.updateProgress(80 + (i / total) * 15, `이미지 OCR (${i + 1}/${total})...`);
            });

            if (ocrCharCount > 0) {
                results.items.push({
                    name: '이미지 내 텍스트',
                    textCharCount: 0,
                    ocrText,
                    ocrCharCount,
                    totalCharCount: ocrCharCount,
                    imageCount: images.length
                });

                results.totalOcrChars += ocrCharCount;
            }
        }

        results.totalChars = results.totalTextChars + results.totalOcrChars;
        this.updateProgress(100, '분석 완료!');
        return results;
    }

    parseSharedStrings(xmlContent) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlContent, 'text/xml');
        const strings = [];
        const siElements = doc.getElementsByTagName('si');

        for (const si of siElements) {
            const tElements = si.getElementsByTagName('t');
            let text = '';
            for (const t of tElements) {
                text += t.textContent || '';
            }
            strings.push(text);
        }

        return strings;
    }

    async getSheetNames(zip) {
        const names = [];
        if (zip.files['xl/workbook.xml']) {
            const wbXml = await zip.files['xl/workbook.xml'].async('text');
            const parser = new DOMParser();
            const doc = parser.parseFromString(wbXml, 'text/xml');
            const sheets = doc.getElementsByTagName('sheet');
            for (const sheet of sheets) {
                names.push(sheet.getAttribute('name') || '');
            }
        }
        return names;
    }

    extractXLSXText(xmlContent, sharedStrings) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlContent, 'text/xml');
        const texts = [];
        const cells = doc.getElementsByTagName('c');

        for (const cell of cells) {
            const type = cell.getAttribute('t');
            const vElement = cell.getElementsByTagName('v')[0];

            if (vElement) {
                if (type === 's') {
                    const index = parseInt(vElement.textContent);
                    if (sharedStrings[index]) {
                        texts.push(sharedStrings[index]);
                    }
                } else if (type === 'inlineStr') {
                    const tElement = cell.getElementsByTagName('t')[0];
                    if (tElement) {
                        texts.push(tElement.textContent);
                    }
                } else if (!type || type === 'n') {
                    texts.push(vElement.textContent);
                }
            }
        }

        return texts.join(' ');
    }

    // ============ PDF 분석 ============
    async analyzePDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        this.updateProgress(10, 'PDF 분석 중...');

        const results = {
            items: [],
            totalTextChars: 0,
            totalOcrChars: 0,
            totalChars: 0
        };

        const totalPages = pdf.numPages;

        for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
            this.updateProgress(
                10 + (pageNum / totalPages) * 80,
                `페이지 ${pageNum}/${totalPages} 분석 중...`
            );

            const page = await pdf.getPage(pageNum);

            // 텍스트 추출
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join('');
            const textCharCount = this.countChars(pageText);

            // 이미지 추출
            const images = await this.extractPDFPageImages(page);
            let ocrText = '';
            let ocrCharCount = 0;

            if (images.length > 0) {
                const ocrResult = await this.processImagesOCR(images, (i, total) => {
                    this.updateProgress(
                        10 + (pageNum / totalPages) * 80,
                        `페이지 ${pageNum} - 이미지 OCR (${i + 1}/${total})...`
                    );
                });
                ocrText = ocrResult.ocrText;
                ocrCharCount = ocrResult.ocrCharCount;
            }

            results.items.push({
                name: `페이지 ${pageNum}`,
                textCharCount,
                ocrText,
                ocrCharCount,
                totalCharCount: textCharCount + ocrCharCount,
                imageCount: images.length
            });

            results.totalTextChars += textCharCount;
            results.totalOcrChars += ocrCharCount;
        }

        results.totalChars = results.totalTextChars + results.totalOcrChars;
        this.updateProgress(100, '분석 완료!');
        return results;
    }

    async extractPDFPageImages(page) {
        const images = [];
        const ops = await page.getOperatorList();

        // 이미지 이름 수집
        const imageNames = new Set();
        for (let i = 0; i < ops.fnArray.length; i++) {
            if (ops.fnArray[i] === pdfjsLib.OPS.paintImageXObject ||
                ops.fnArray[i] === pdfjsLib.OPS.paintJpegXObject) {
                imageNames.add(ops.argsArray[i][0]);
            }
        }

        // 이미지가 없으면 바로 반환
        if (imageNames.size === 0) {
            return images;
        }

        // 페이지 렌더링으로 이미지 객체 로드 트리거
        const scale = this.getOptimalPDFScale(page);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');

        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        // 렌더링 실행 (이미지 객체 로드됨)
        await page.render({
            canvasContext: ctx,
            viewport: viewport
        }).promise;

        // 각 이미지 추출 시도
        for (const imgName of imageNames) {
            try {
                const imgData = await this.getPDFImageData(page, imgName);
                if (imgData) {
                    const dataUrl = this.convertPDFImageToDataURL(imgData);
                    if (dataUrl) {
                        images.push(dataUrl);
                    }
                }
            } catch (err) {
                // 개별 이미지 추출 실패 시 무시
            }
        }

        // 개별 이미지 추출 실패 시 폴백: 페이지 전체 렌더링 사용
        if (images.length === 0 && imageNames.size > 0) {
            images.push(canvas.toDataURL('image/png'));
        }

        return images;
    }

    async getPDFImageData(page, imgName) {
        return new Promise((resolve, reject) => {
            let attempts = 0;
            const maxAttempts = 40; // 2초 타임아웃

            const checkObj = () => {
                if (page.objs.has(imgName)) {
                    resolve(page.objs.get(imgName));
                } else if (page.commonObjs.has(imgName)) {
                    resolve(page.commonObjs.get(imgName));
                } else if (attempts++ < maxAttempts) {
                    setTimeout(checkObj, 50);
                } else {
                    reject(new Error('이미지 로드 타임아웃'));
                }
            };
            checkObj();
        });
    }

    convertPDFImageToDataURL(imgData) {
        if (!imgData) {
            return null;
        }

        const { width, height, data, bitmap } = imgData;

        // 매우 작은 이미지는 제외 (10px 이하)
        if (width <= 10 || height <= 10) {
            return null;
        }

        // 원본 크기 그대로 추출 (스케일링은 normalizeImageForOCR에서 통일)
        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');

        // ImageBitmap이 있으면 직접 그리기
        if (bitmap) {
            ctx.drawImage(bitmap, 0, 0, width, height);
            return canvas.toDataURL('image/png');
        }

        // data 배열이 있으면 기존 방식으로 처리
        if (!data) {
            return null;
        }

        // 원본 크기로 이미지 데이터 생성 (스케일링은 normalizeImageForOCR에서 통일)
        const imageData = ctx.createImageData(width, height);
        const srcData = data;
        const dstData = imageData.data;
        const pixelCount = width * height;

        if (srcData.length === pixelCount * 4) {
            dstData.set(srcData);
        } else if (srcData.length === pixelCount * 3) {
            for (let j = 0, k = 0; j < srcData.length; j += 3, k += 4) {
                dstData[k] = srcData[j];
                dstData[k + 1] = srcData[j + 1];
                dstData[k + 2] = srcData[j + 2];
                dstData[k + 3] = 255;
            }
        } else if (srcData.length === pixelCount) {
            for (let j = 0, k = 0; j < srcData.length; j++, k += 4) {
                dstData[k] = srcData[j];
                dstData[k + 1] = srcData[j];
                dstData[k + 2] = srcData[j];
                dstData[k + 3] = 255;
            }
        } else {
            return null;
        }

        ctx.putImageData(imageData, 0, 0);
        return canvas.toDataURL('image/png');
    }

    async renderPDFPageToImage(page) {
        // 동적 스케일 계산 (300 DPI 목표)
        const scale = this.getOptimalPDFScale(page);
        const viewport = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');

        // 흰색 배경 설정 (투명 배경 문제 방지)
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);

        await page.render({
            canvasContext: ctx,
            viewport: viewport
        }).promise;

        return canvas.toDataURL('image/png');
    }

    // ============ 공통 유틸리티 ============
    getSortedFiles(zip, pattern) {
        const files = [];
        for (const filename of Object.keys(zip.files)) {
            const match = filename.match(pattern);
            if (match) {
                files.push({ filename, num: parseInt(match[1]) });
            }
        }
        files.sort((a, b) => a.num - b.num);
        return files.map(f => f.filename);
    }

    async extractImagesFromZip(zip, mediaFolder) {
        const images = [];

        for (const filename of Object.keys(zip.files)) {
            if (filename.startsWith(mediaFolder)) {
                const ext = filename.toLowerCase().split('.').pop();
                if (this.imageFormats.includes(ext)) {
                    const data = await zip.files[filename].async('base64');
                    const mimeType = this.getMimeType(ext);
                    images.push({
                        filename,
                        name: filename.replace(mediaFolder, ''),
                        data: `data:${mimeType};base64,${data}`,
                        type: mimeType
                    });
                }
            }
        }

        return images;
    }

    getMimeType(ext) {
        const types = {
            'png': 'image/png',
            'jpg': 'image/jpeg',
            'jpeg': 'image/jpeg',
            'gif': 'image/gif',
            'bmp': 'image/bmp',
            'tiff': 'image/tiff',
            'webp': 'image/webp'
        };
        return types[ext] || 'image/png';
    }

    extractTextFromXML(xmlContent, tagName) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xmlContent, 'text/xml');
        const textElements = doc.getElementsByTagName(tagName);
        const texts = [];

        for (const element of textElements) {
            if (element.textContent) {
                texts.push(element.textContent);
            }
        }

        // 텍스트를 공백 없이 연결 (아랍어 등 RTL 언어 지원)
        // 각 텍스트 조각 사이에 공백 대신 빈 문자열로 연결하고,
        // 연속된 공백은 하나로 정리
        return texts.join('').replace(/\s+/g, ' ').trim();
    }

    countChars(text) {
        return text.replace(/\s/g, '').length;
    }

    updateProgress(percent, text) {
        this.progressFill.style.width = `${percent}%`;
        this.progressText.textContent = text;
    }

    displayResults(results) {
        this.progressSection.classList.remove('show');
        this.resultsSection.classList.add('show');

        document.getElementById('totalChars').textContent = results.totalChars.toLocaleString();
        document.getElementById('textChars').textContent = results.totalTextChars.toLocaleString();
        document.getElementById('ocrChars').textContent = results.totalOcrChars.toLocaleString();

        this.slidesContainer.innerHTML = '';

        for (const item of results.items) {
            const card = document.createElement('div');
            card.className = 'detail-card';

            card.innerHTML = `
                <div class="detail-header">
                    <span class="detail-title">${item.name}</span>
                    <span class="detail-count">${item.totalCharCount.toLocaleString()}자</span>
                </div>
                <div class="detail-body">
                    <div class="detail-row">
                        <span class="detail-label">텍스트 글자수</span>
                        <span class="detail-value">${item.textCharCount.toLocaleString()}자</span>
                    </div>
                    ${item.imageCount > 0 ? `
                        <div class="detail-row">
                            <span class="detail-label">이미지 수</span>
                            <span class="detail-value">${item.imageCount}개</span>
                        </div>
                    ` : ''}
                    ${item.ocrCharCount > 0 ? `
                        <div class="detail-row">
                            <span class="detail-label">이미지 내 글자수 (OCR)</span>
                            <span class="detail-value">${item.ocrCharCount.toLocaleString()}자</span>
                        </div>
                    ` : ''}
                    ${item.ocrText ? `
                        <div class="ocr-preview">
                            <h4>OCR로 인식된 텍스트:</h4>
                            <p dir="auto">${this.escapeHtml(item.ocrText)}</p>
                        </div>
                    ` : ''}
                </div>
            `;

            this.slidesContainer.appendChild(card);
        }
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    getImageSize(dataUrl) {
        return new Promise((resolve) => {
            const img = new Image();
            img.onload = () => resolve({ width: img.width, height: img.height });
            img.onerror = () => resolve({ width: 0, height: 0 });
            img.src = dataUrl;
        });
    }

    showError(message) {
        this.errorMessage.textContent = message;
        this.errorMessage.classList.add('show');
    }

    hideError() {
        this.errorMessage.classList.remove('show');
    }
}

// 앱 초기화
document.addEventListener('DOMContentLoaded', () => {
    new DocumentTextCounter();
});
