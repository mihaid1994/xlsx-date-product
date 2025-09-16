// js/basefunc.js
class ExcelProcessor {
  constructor() {
    this.requiredColumns = ["Дата изготовления", "Срок годности"];
    this.newColumns = ["Срок годности в месяцах общий", "Осталось месяцев"];
    this.uploadedFiles = [];
    this.processedFiles = [];
    this.init();
  }

  init() {
    this.setupEventListeners();
    this.updateCurrentDate();
  }

  setupEventListeners() {
    const fileInput = document.getElementById("fileInput");
    const uploadArea = document.getElementById("uploadArea");
    const processBtn = document.getElementById("processBtn");

    // Обработка выбора файлов
    fileInput.addEventListener("change", (e) =>
      this.handleFiles(e.target.files)
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

    uploadArea.addEventListener("drop", (e) => {
      e.preventDefault();
      uploadArea.classList.remove("dragover");
      this.handleFiles(e.dataTransfer.files);
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

  handleFiles(files) {
    this.uploadedFiles = [];

    for (let file of files) {
      if (this.isValidExcelFile(file)) {
        this.uploadedFiles.push(file);
      } else {
        this.showError(`Неподдерживаемый формат файла: ${file.name}`);
      }
    }

    if (this.uploadedFiles.length > 0) {
      this.displayUploadedFiles();
    }
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
      const fileItem = document.createElement("div");
      fileItem.className = "file-item fade-in";
      fileItem.innerHTML = `
                <div class="file-info">
                    <i class="fas fa-file-excel file-icon"></i>
                    <div class="file-details">
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${(file.size / 1024).toFixed(
                          1
                        )} KB</div>
                    </div>
                </div>
            `;
      filesList.appendChild(fileItem);
    });

    uploadedFilesDiv.style.display = "block";
    uploadedFilesDiv.classList.add("fade-in");
  }

  async processFiles() {
    const processBtn = document.getElementById("processBtn");
    const progressSection = document.getElementById("progressSection");
    const resultsSection = document.getElementById("resultsSection");

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

          // Преобразуем в JSON
          const jsonData = XLSX.utils.sheet_to_json(worksheet);

          if (jsonData.length === 0) {
            reject(new Error("Файл пустой или не содержит данных"));
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

  findAvailableColumns(existingColumns, newColumns) {
    // Просто возвращаем оригинальные имена - будем добавлять в конец
    const availableColumnNames = {};

    newColumns.forEach((newColumnName) => {
      availableColumnNames[newColumnName] = newColumnName;
    });

    return availableColumnNames;
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

    // Находим доступные имена для новых колонок (просто оригинальные имена)
    const columnMapping = this.findAvailableColumns(
      [...availableColumns],
      [...this.newColumns]
    );

    // Обрабатываем каждую строку
    const processedData = data.map((row) => {
      const newRow = { ...row };

      // Парсим даты
      const manufactureDate = this.parseDate(row["Дата изготовления"]);
      const expiryDate = this.parseDate(row["Срок годности"]);

      // Рассчитываем общий срок годности в месяцах
      const totalMonthsColumnName =
        columnMapping["Срок годности в месяцах общий"];
      newRow[totalMonthsColumnName] = this.calculateMonthsDifference(
        manufactureDate,
        expiryDate
      );

      // Рассчитываем оставшиеся месяцы
      const currentDate = new Date();
      const remainingMonthsColumnName = columnMapping["Осталось месяцев"];
      newRow[remainingMonthsColumnName] = this.calculateMonthsDifference(
        currentDate,
        expiryDate
      );

      return newRow;
    });

    return processedData;
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
            const month = parseInt(parts[1]) - 1; // месяцы в JS начинаются с 0
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
    // Можно добавить toast уведомления
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

  console.log("Excel Processor инициализирован");
});
