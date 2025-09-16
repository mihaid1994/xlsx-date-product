// js/basefunc.js - Улучшенная версия с надежным парсингом Excel
class ExcelProcessor {
  constructor() {
    this.requiredColumns = ["Дата изготовления", "Срок годности"];
    this.newColumns = ["Срок годности в месяцах общий", "Осталось месяцев"];
    this.uploadedFiles = [];
    this.processedFiles = [];
    this.fileAnalysis = [];
    this.init();
  }

  init() {
    this.setupEventListeners();
    this.updateCurrentDate();
    this.injectAnalysisStyles();
  }

  injectAnalysisStyles() {
    const style = document.createElement("style");
    style.textContent = `
      .file-analysis {
        margin-top: 15px;
        padding: 15px;
        background: #f8f9fa;
        border-radius: 8px;
        border-left: 4px solid #007bff;
      }
      .analysis-header {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 12px;
        font-size: 14px;
      }
      .status-success { color: #28a745; }
      .status-error { color: #dc3545; }
      .analysis-details {
        font-size: 13px;
        line-height: 1.5;
      }
      .detail-row {
        margin-bottom: 8px;
        display: flex;
        flex-wrap: wrap;
        align-items: flex-start;
        gap: 8px;
      }
      .detail-row span:first-child {
        color: #666;
        min-width: 200px;
      }
      .new-columns-info {
        flex-direction: column;
        align-items: flex-start;
        background: #e8f4fd;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
      }
      .new-columns-list {
        margin-top: 5px;
        width: 100%;
      }
      .new-column {
        color: #0066cc;
        font-size: 12px;
        margin-bottom: 3px;
      }
      .error-row {
        background: #fff2f2;
        padding: 8px;
        border-radius: 4px;
        border-left: 3px solid #dc3545;
        flex-direction: column;
        align-items: flex-start;
      }
      .missing-columns {
        color: #dc3545;
        font-weight: bold;
        margin-top: 5px;
      }
      .error-message {
        color: #dc3545;
        font-weight: bold;
        margin-top: 8px;
      }
    `;
    document.head.appendChild(style);
  }

  setupEventListeners() {
    const fileInput = document.getElementById("fileInput");
    const uploadArea = document.getElementById("uploadArea");
    const processBtn = document.getElementById("processBtn");

    // Обработка выбора файлов
    fileInput.addEventListener(
      "change",
      async (e) => await this.handleFiles(e.target.files)
    );

    // Drag & Drop
    uploadArea.addEventListener("dragover", (e) => {
      e.preventDefault();
      uploadArea.classList.add("dragover");
    });

    uploadArea.addEventListener("dragleave", (e) => {
      e.preventDefault();
      uploadArea.classList.remove("dragover");
    });

    uploadArea.addEventListener("drop", async (e) => {
      e.preventDefault();
      uploadArea.classList.remove("dragover");
      await this.handleFiles(e.dataTransfer.files);
    });

    // Обработка файлов
    processBtn.addEventListener("click", () => this.processFiles());
  }

  updateCurrentDate() {
    const now = new Date();
    const dateStr = now.toLocaleDateString("ru-RU", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
    });
    document.getElementById("currentDate").textContent = dateStr;
  }

  async handleFiles(files) {
    this.uploadedFiles = [];

    for (let file of files) {
      if (this.isValidExcelFile(file)) {
        this.uploadedFiles.push(file);
      } else {
        this.showError(`Неподдерживаемый формат файла: ${file.name}`);
      }
    }

    if (this.uploadedFiles.length > 0) {
      await this.analyzeFilesBeforeProcessing();
      this.displayUploadedFiles();
    }
  }

  async analyzeFilesBeforeProcessing() {
    const analysisResults = [];

    for (let file of this.uploadedFiles) {
      try {
        const analysis = await this.analyzeFile(file);
        analysisResults.push(analysis);
      } catch (error) {
        console.error(`Ошибка анализа файла ${file.name}:`, error);
        analysisResults.push({
          fileName: file.name,
          error: error.message,
        });
      }
    }

    this.fileAnalysis = analysisResults;
  }

  // УЛУЧШЕННЫЙ МЕТОД АНАЛИЗА ФАЙЛА
  async analyzeFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];

          // НОВЫЙ ПОДХОД: Анализ структуры файла напрямую из worksheet
          const structureInfo = this.analyzeWorksheetStructure(worksheet);

          if (structureInfo.totalRows === 0) {
            reject(new Error("Файл пустой или не содержит данных"));
            return;
          }

          // Получаем заголовки из первой строки
          const headers = this.extractHeaders(
            worksheet,
            structureInfo.dataRange
          );

          // Проверяем наличие обязательных колонок
          const missingColumns = this.requiredColumns.filter(
            (col) => !headers.includes(col)
          );

          // Определяем позиции для новых столбцов
          const newColumnPositions = this.calculateNewColumnPositions(
            structureInfo.lastColumn
          );

          resolve({
            fileName: file.name,
            totalColumns: structureInfo.lastColumn + 1,
            totalRows: structureInfo.totalRows - 1, // Исключаем заголовок
            dataRange: structureInfo.dataRange,
            existingColumns: headers,
            lastFilledColumn: this.columnIndexToLetter(
              structureInfo.lastColumn
            ),
            newColumnsWillBe: this.newColumns.map((col, index) => ({
              name: col,
              letter: this.columnIndexToLetter(
                newColumnPositions.startIndex + index
              ),
              position: newColumnPositions.startIndex + index + 1,
            })),
            missingRequiredColumns: missingColumns,
            isValid: missingColumns.length === 0,
          });
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error("Ошибка чтения файла"));
      reader.readAsArrayBuffer(file);
    });
  }

  // НОВЫЙ МЕТОД: Анализ структуры worksheet напрямую
  analyzeWorksheetStructure(worksheet) {
    if (!worksheet || !worksheet["!ref"]) {
      return { lastColumn: -1, totalRows: 0, dataRange: null };
    }

    const range = XLSX.utils.decode_range(worksheet["!ref"]);

    // Найдем последний заполненный столбец более надежным способом
    let lastFilledColumn = -1;
    let totalDataRows = 0;

    // Проходим по всем ячейкам для точного определения структуры
    for (let row = range.s.r; row <= range.e.r; row++) {
      let hasDataInRow = false;

      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];

        // Проверяем, есть ли данные в ячейке
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
          lastFilledColumn = Math.max(lastFilledColumn, col);
          hasDataInRow = true;
        }
      }

      if (hasDataInRow) {
        totalDataRows++;
      }
    }

    return {
      lastColumn: lastFilledColumn,
      totalRows: totalDataRows,
      dataRange: range,
      startRow: range.s.r,
      startCol: range.s.c,
      endRow: range.e.r,
      endCol: range.e.c,
    };
  }

  // НОВЫЙ МЕТОД: Извлечение заголовков напрямую из первой строки
  extractHeaders(worksheet, range) {
    if (!range) return [];

    const headers = [];
    const headerRow = range.s.r; // Первая строка

    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: headerRow, c: col });
      const cell = worksheet[cellAddress];

      if (cell && cell.v !== undefined && cell.v !== null && cell.v !== "") {
        headers.push(String(cell.v).trim());
      } else {
        // Для пустых ячеек добавляем placeholder, чтобы сохранить позиции
        headers.push(`Столбец_${col + 1}`);
      }
    }

    return headers;
  }

  // УПРОЩЕННЫЙ МЕТОД: Расчет позиций новых столбцов
  calculateNewColumnPositions(lastColumnIndex) {
    return {
      startIndex: lastColumnIndex + 1,
      endIndex: lastColumnIndex + this.newColumns.length,
    };
  }

  // ВСПОМОГАТЕЛЬНЫЙ МЕТОД: Конвертация индекса столбца в букву (A, B, C...)
  columnIndexToLetter(index) {
    if (index < 0) return "";

    let letter = "";
    while (index >= 0) {
      letter = String.fromCharCode(65 + (index % 26)) + letter;
      index = Math.floor(index / 26) - 1;
    }
    return letter;
  }

  isValidExcelFile(file) {
    const validExtensions = [".xlsx", ".xls"];
    const fileName = file.name.toLowerCase();
    return validExtensions.some((ext) => fileName.endsWith(ext));
  }

  displayUploadedFiles() {
    const uploadedFilesDiv = document.getElementById("uploadedFiles");
    const filesList = document.getElementById("filesList");

    filesList.innerHTML = "";

    this.uploadedFiles.forEach((file, index) => {
      const analysis = this.fileAnalysis[index];

      const fileItem = document.createElement("div");
      fileItem.className = "file-item fade-in";

      let analysisHtml = "";
      if (analysis && !analysis.error) {
        const statusClass = analysis.isValid
          ? "status-success"
          : "status-error";
        const statusIcon = analysis.isValid
          ? "fa-check-circle"
          : "fa-exclamation-triangle";

        analysisHtml = `
          <div class="file-analysis">
            <div class="analysis-header">
              <i class="fas ${statusIcon} ${statusClass}"></i>
              <strong>Анализ файла:</strong>
            </div>
            <div class="analysis-details">
              <div class="detail-row">
                <span>📊 Строк данных:</span> 
                <strong>${analysis.totalRows}</strong>
              </div>
              <div class="detail-row">
                <span>📋 Столбцов:</span> 
                <strong>${analysis.totalColumns}</strong>
              </div>
              <div class="detail-row">
                <span>🎯 Последний заполненный столбец:</span> 
                <strong>"${analysis.lastFilledColumn}"</strong>
              </div>
              <div class="detail-row">
                <span>📝 Диапазон данных:</span> 
                <strong>${
                  analysis.dataRange
                    ? XLSX.utils.encode_range(analysis.dataRange)
                    : "Не определен"
                }</strong>
              </div>
              <div class="detail-row new-columns-info">
                <span>➕ Новые столбцы будут добавлены:</span>
                <div class="new-columns-list">
                  ${analysis.newColumnsWillBe
                    .map(
                      (col) =>
                        `<div class="new-column">• Столбец ${col.letter} (позиция ${col.position}): "<strong>${col.name}</strong>"</div>`
                    )
                    .join("")}
                </div>
              </div>
              ${
                analysis.missingRequiredColumns.length > 0
                  ? `
                <div class="detail-row error-row">
                  <span>❌ Отсутствуют обязательные колонки:</span>
                  <div class="missing-columns">${analysis.missingRequiredColumns.join(
                    ", "
                  )}</div>
                </div>
              `
                  : ""
              }
            </div>
          </div>
        `;
      } else if (analysis && analysis.error) {
        analysisHtml = `
          <div class="file-analysis">
            <div class="analysis-header">
              <i class="fas fa-exclamation-triangle status-error"></i>
              <strong>Ошибка анализа:</strong>
            </div>
            <div class="error-message">${analysis.error}</div>
          </div>
        `;
      }

      fileItem.innerHTML = `
        <div class="file-info">
          <i class="fas fa-file-excel file-icon"></i>
          <div class="file-details">
            <div class="file-name">${file.name}</div>
            <div class="file-size">${(file.size / 1024).toFixed(1)} KB</div>
          </div>
        </div>
        ${analysisHtml}
      `;

      filesList.appendChild(fileItem);
    });

    uploadedFilesDiv.style.display = "block";
    uploadedFilesDiv.classList.add("fade-in");

    // Добавляем общую информацию о готовности
    this.displayProcessingReadiness();
  }

  displayProcessingReadiness() {
    let existingReadiness = document.getElementById("processingReadiness");
    if (existingReadiness) {
      existingReadiness.remove();
    }

    const validFiles = this.fileAnalysis.filter(
      (analysis) => !analysis.error && analysis.isValid
    ).length;

    const invalidFiles = this.fileAnalysis.length - validFiles;

    const readinessDiv = document.createElement("div");
    readinessDiv.id = "processingReadiness";
    readinessDiv.className = "processing-readiness fade-in";

    if (invalidFiles === 0) {
      readinessDiv.innerHTML = `
        <div class="readiness-success">
          <i class="fas fa-check-circle"></i>
          <strong>Готово к обработке!</strong> Все ${validFiles} файлов прошли проверку.
        </div>
      `;
    } else {
      readinessDiv.innerHTML = `
        <div class="readiness-warning">
          <i class="fas fa-exclamation-triangle"></i>
          <strong>Внимание:</strong> ${validFiles} файлов готово к обработке, ${invalidFiles} с ошибками.
          <br><small>Файлы с ошибками не будут обрабатываться.</small>
        </div>
      `;
    }

    // Добавляем стили для этого блока
    if (!document.getElementById("readiness-styles")) {
      const style = document.createElement("style");
      style.id = "readiness-styles";
      style.textContent = `
        .processing-readiness {
          margin: 15px 0;
          padding: 15px;
          border-radius: 8px;
          text-align: center;
        }
        .readiness-success {
          background: #d4edda;
          color: #155724;
          border: 1px solid #c3e6cb;
        }
        .readiness-warning {
          background: #fff3cd;
          color: #856404;
          border: 1px solid #ffeaa7;
        }
        .readiness-success i, .readiness-warning i {
          margin-right: 8px;
        }
      `;
      document.head.appendChild(style);
    }

    const uploadedFilesDiv = document.getElementById("uploadedFiles");
    uploadedFilesDiv.appendChild(readinessDiv);
  }

  async processFiles() {
    const processBtn = document.getElementById("processBtn");
    const progressSection = document.getElementById("progressSection");
    const resultsSection = document.getElementById("resultsSection");

    // Проверяем, есть ли файлы с ошибками анализа
    const invalidFiles = this.fileAnalysis.filter(
      (analysis) => analysis.error || !analysis.isValid
    );

    if (invalidFiles.length > 0) {
      const fileNames = invalidFiles.map((f) => f.fileName).join(", ");
      alert(
        `Невозможно обработать файлы с ошибками: ${fileNames}\nПроверьте структуру файлов и повторите попытку.`
      );
      return;
    }

    // Подготовка
    processBtn.disabled = true;
    progressSection.style.display = "block";
    progressSection.classList.add("fade-in");
    resultsSection.style.display = "none";

    this.processedFiles = [];
    const errors = [];

    for (let i = 0; i < this.uploadedFiles.length; i++) {
      const file = this.uploadedFiles[i];

      try {
        this.updateProgress(
          (i / this.uploadedFiles.length) * 100,
          `Обработка файла: ${file.name}`
        );

        const processedData = await this.processFile(file);
        this.processedFiles.push({
          name: `processed_${file.name.replace(/\.(xls|xlsx)$/i, ".xlsx")}`,
          originalName: file.name,
          data: processedData,
        });
      } catch (error) {
        console.error(`Ошибка обработки файла ${file.name}:`, error);
        errors.push(`Ошибка в файле ${file.name}: ${error.message}`);
      }
    }

    this.updateProgress(100, "Обработка завершена!");

    // Показываем результаты
    setTimeout(() => {
      progressSection.style.display = "none";
      this.displayResults(errors);
      processBtn.disabled = false;
    }, 1000);
  }

  // УЛУЧШЕННЫЙ МЕТОД ОБРАБОТКИ ФАЙЛА
  async processFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          // Берем первый лист
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];

          // НОВЫЙ ПОДХОД: Используем улучшенную структуру
          const structureInfo = this.analyzeWorksheetStructure(worksheet);

          if (structureInfo.totalRows <= 1) {
            // <= 1 потому что первая строка - заголовки
            reject(new Error("Файл пустой или содержит только заголовки"));
            return;
          }

          // Конвертируем в JSON с настройками для корректного парсинга
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            raw: false, // Не использовать raw значения
            dateNF: "dd.mm.yyyy", // Формат даты
            defval: "", // Значение по умолчанию для пустых ячеек
          });

          if (jsonData.length === 0) {
            reject(new Error("Не удалось извлечь данные из файла"));
            return;
          }

          // Обрабатываем данные
          const processedData = this.processDataFrame(jsonData);

          // Создаем новый Excel файл
          const newWorkbook = XLSX.utils.book_new();
          const newWorksheet = XLSX.utils.json_to_sheet(processedData);

          // Устанавливаем ширину столбцов
          this.setColumnWidths(newWorksheet, processedData);

          XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "TDSheet");

          // Генерируем файл в формате .xlsx для совместимости
          const excelBuffer = XLSX.write(newWorkbook, {
            bookType: "xlsx",
            type: "array",
          });

          resolve(excelBuffer);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error("Ошибка чтения файла"));
      reader.readAsArrayBuffer(file);
    });
  }

  // УПРОЩЕННЫЙ МЕТОД ОБРАБОТКИ ДАННЫХ
  processDataFrame(data) {
    // Проверяем наличие обязательных колонок
    const firstRow = data[0];
    const availableColumns = Object.keys(firstRow);

    const missingColumns = this.requiredColumns.filter(
      (col) => !availableColumns.includes(col)
    );

    if (missingColumns.length > 0) {
      throw new Error(
        `Отсутствуют обязательные колонки: ${missingColumns.join(", ")}`
      );
    }

    // Обрабатываем каждую строку - просто добавляем новые столбцы в конец
    const processedData = data.map((row) => {
      const newRow = { ...row };

      // Парсим даты
      const manufactureDate = this.parseDate(row["Дата изготовления"]);
      const expiryDate = this.parseDate(row["Срок годности"]);

      // Добавляем новые столбцы
      newRow["Срок годности в месяцах общий"] = this.calculateMonthsDifference(
        manufactureDate,
        expiryDate
      );

      newRow["Осталось месяцев"] = this.calculateMonthsDifference(
        new Date(),
        expiryDate
      );

      return newRow;
    });

    return processedData;
  }

  setColumnWidths(worksheet, data) {
    const cols = [];

    if (data.length === 0) return;

    const headers = Object.keys(data[0]);

    headers.forEach((header, index) => {
      let width = 15; // Минимальная ширина по умолчанию

      // Особые случаи для определенных столбцов
      if (
        header.toLowerCase().includes("наименование") ||
        header.toLowerCase().includes("название") ||
        header.toLowerCase().includes("товар") ||
        header.toLowerCase().includes("продукт")
      ) {
        width = 35; // Больше для наименований
      } else if (header.includes("Дата") || header.includes("Срок")) {
        width = 20; // Для дат
      } else if (header.includes("месяц") || header.includes("Осталось")) {
        width = 18; // Для расчетных полей
      }

      // Проверяем максимальную длину содержимого в столбце
      const maxContentLength = Math.max(
        header.length,
        ...data.slice(0, 100).map((row) => {
          // Проверяем первые 100 строк для производительности
          const value = row[header];
          return value ? String(value).length : 0;
        })
      );

      // Устанавливаем ширину как максимум между минимальной и длиной содержимого
      width = Math.max(width, Math.min(maxContentLength + 2, 50)); // +2 для отступов, максимум 50

      cols.push({ width: width });
    });

    worksheet["!cols"] = cols;
  }

  parseDate(dateStr) {
    if (!dateStr || dateStr === "" || dateStr == null) {
      return null;
    }

    try {
      // Если это уже объект Date
      if (dateStr instanceof Date) {
        return dateStr;
      }

      // Если это строка
      if (typeof dateStr === "string") {
        dateStr = dateStr.trim();

        // Формат DD.MM.YYYY
        if (dateStr.includes(".")) {
          const parts = dateStr.split(".");
          if (parts.length === 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1; // месяцы в JS начинают с 0
            const year = parseInt(parts[2]);
            return new Date(year, month, day);
          }
        }

        // Формат DD/MM/YYYY
        if (dateStr.includes("/")) {
          const parts = dateStr.split("/");
          if (parts.length === 3) {
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const year = parseInt(parts[2]);
            return new Date(year, month, day);
          }
        }

        // Формат YYYY-MM-DD
        if (dateStr.includes("-")) {
          return new Date(dateStr);
        }
      }

      // Если это число (Excel serial date)
      if (typeof dateStr === "number") {
        // Excel даты начинаются с 1900-01-01
        const excelEpoch = new Date(1900, 0, 1);
        const msPerDay = 24 * 60 * 60 * 1000;
        return new Date(excelEpoch.getTime() + (dateStr - 2) * msPerDay);
      }

      return null;
    } catch (error) {
      console.warn(`Не удалось распарсить дату: ${dateStr}`, error);
      return null;
    }
  }

  calculateMonthsDifference(startDate, endDate) {
    if (!startDate || !endDate) {
      return null;
    }

    try {
      // Разность в годах * 12 + разность в месяцах
      const months =
        (endDate.getFullYear() - startDate.getFullYear()) * 12 +
        (endDate.getMonth() - startDate.getMonth());

      // Учитываем дни для более точного расчета
      if (endDate.getDate() < startDate.getDate()) {
        return months - 1;
      }

      return months;
    } catch (error) {
      console.error("Ошибка расчета месяцев:", error);
      return null;
    }
  }

  updateProgress(percentage, text) {
    const progressFill = document.getElementById("progressFill");
    const progressText = document.getElementById("progressText");

    progressFill.style.width = percentage + "%";
    progressText.textContent = text;
  }

  displayResults(errors) {
    const resultsSection = document.getElementById("resultsSection");
    const resultsContent = document.getElementById("resultsContent");

    resultsContent.innerHTML = "";

    // Успешные файлы
    if (this.processedFiles.length > 0) {
      const successDiv = document.createElement("div");
      successDiv.className = "success-message";
      successDiv.innerHTML = `
                <i class="fas fa-check-circle"></i>
                Успешно обработано файлов: ${this.processedFiles.length}
            `;
      resultsContent.appendChild(successDiv);

      const downloadSection = document.createElement("div");
      downloadSection.className = "download-section";

      // Кнопка скачать все файлы (ZIP)
      if (this.processedFiles.length > 1) {
        const downloadAllBtn = document.createElement("button");
        downloadAllBtn.className = "download-all-btn";
        downloadAllBtn.innerHTML = `
                    <i class="fas fa-download"></i>
                    Скачать все файлы (ZIP)
                `;
        downloadAllBtn.addEventListener("click", () => this.downloadAllFiles());
        downloadSection.appendChild(downloadAllBtn);
      }

      // Отдельные кнопки для каждого файла
      const individualDiv = document.createElement("div");
      individualDiv.className = "individual-downloads";
      individualDiv.innerHTML = "<h4>📄 Скачать отдельные файлы:</h4>";

      this.processedFiles.forEach((file) => {
        const downloadItem = document.createElement("div");
        downloadItem.className = "download-item";
        downloadItem.innerHTML = `
                    <div>
                        <i class="fas fa-file-excel" style="color: #28a745; margin-right: 10px;"></i>
                        <strong>${file.name}</strong>
                        <br>
                        <small style="color: #666;">Исходный: ${file.originalName}</small>
                    </div>
                    <button class="download-btn" onclick="excelProcessor.downloadFile('${file.name}')">
                        <i class="fas fa-download"></i>
                        Скачать
                    </button>
                `;
        individualDiv.appendChild(downloadItem);
      });

      downloadSection.appendChild(individualDiv);
      resultsContent.appendChild(downloadSection);
    }

    // Ошибки
    if (errors.length > 0) {
      const errorDiv = document.createElement("div");
      errorDiv.className = "error-section";
      errorDiv.innerHTML = `<h4><i class="fas fa-exclamation-triangle"></i> Ошибки в ${errors.length} файлах:</h4>`;

      errors.forEach((error) => {
        const errorItem = document.createElement("div");
        errorItem.className = "error-item";
        errorItem.textContent = error;
        errorDiv.appendChild(errorItem);
      });

      resultsContent.appendChild(errorDiv);
    }

    resultsSection.style.display = "block";
    resultsSection.classList.add("fade-in");
  }

  downloadFile(fileName) {
    const file = this.processedFiles.find((f) => f.name === fileName);
    if (file) {
      const blob = new Blob([file.data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(blob, fileName);
    }
  }

  async downloadAllFiles() {
    if (typeof JSZip === "undefined") {
      // Если JSZip недоступен, скачиваем файлы по одному
      this.processedFiles.forEach((file) => this.downloadFile(file.name));
      return;
    }

    const zip = new JSZip();

    this.processedFiles.forEach((file) => {
      zip.file(file.name, file.data);
    });

    const content = await zip.generateAsync({ type: "blob" });
    const timestamp = new Date()
      .toISOString()
      .slice(0, 19)
      .replace(/[:]/g, "-");
    saveAs(content, `processed_files_${timestamp}.zip`);
  }

  showError(message) {
    console.error(message);
    alert(message);
  }
}

// Инициализация приложения
let excelProcessor;

document.addEventListener("DOMContentLoaded", function () {
  excelProcessor = new ExcelProcessor();

  // Проверяем поддержку браузера
  if (typeof XLSX === "undefined") {
    document.body.innerHTML = `
            <div style="padding: 50px; text-align: center; color: red;">
                <h2>Ошибка загрузки библиотек</h2>
                <p>Не удалось загрузить необходимые библиотеки для работы с Excel файлами.</p>
                <p>Проверьте подключение к интернету и перезагрузите страницу.</p>
            </div>
        `;
    return;
  }

  console.log("Excel Processor инициализирован с улучшенным парсингом");
});
