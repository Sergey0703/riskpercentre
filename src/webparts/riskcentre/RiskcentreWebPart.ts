import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import * as ExcelJS from 'exceljs';

import styles from './RiskcentreWebPart.module.scss';

export interface IRiskcentreWebPartProps {
  description: string;
}

interface IExcelFileInfo {
  name: string;
  url: string;
  hasSheet1: boolean;
  headerMatch: boolean;
  errorMessage?: string;
}

export default class RiskcentreWebPart extends BaseClientSideWebPart<IRiskcentreWebPartProps> {
  private sp: SPFI;
  private readonly FOLDER_PATH = "Shared Documents/Risk per centre";
  private readonly SUMMARY_FILE_NAME = "_OVERVIEW SUMMARY.xlsx";
  private readonly SHEET_NAME = "Sheet1";
  private excelFiles: IExcelFileInfo[] = [];
  private summaryFilePath: string = '';
  private expectedHeaders: string[] = [];

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));
    
    // Формируем полный путь к папке
    const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/${this.FOLDER_PATH}`;
    this.summaryFilePath = `${folderUrl}/${this.SUMMARY_FILE_NAME}`;
    
    console.log('[RiskCentre] Initialized:', {
      folderUrl,
      summaryFilePath: this.summaryFilePath
    });
  }

  public async render(): Promise<void> {
    this.domElement.innerHTML = `
      <div class="${styles.riskcentre}">
        <div class="${styles.header}">
          <h2>Risk per Centre - Data Aggregator</h2>
          <p>Combine Excel files from Risk per centre folder into _OVERVIEW SUMMARY.xlsx</p>
        </div>

        <div id="loadingSpinner" class="${styles.spinnerContainer}" style="display: none;">
          <div class="${styles.spinner}"></div>
          <p id="loadingMessage">Loading files...</p>
        </div>

        <div id="errorMessage" class="${styles.errorMessage}" style="display: none;"></div>
        <div id="warningMessage" class="${styles.warningMessage}" style="display: none;"></div>

        <div id="fileTableContainer" style="display: none;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; gap: 20px;">
            <div style="display: flex; align-items: center; gap: 16px; flex: 1;">
              <h3 style="color: #323130; font-size: 18px; font-weight: 600; margin: 0;">Select Files to Process</h3>
              <div style="background-color: #0078d4; color: white; padding: 4px 12px; border-radius: 12px; font-size: 13px; font-weight: 600;">
                <span id="fileCount">0</span> files
              </div>
              <div style="display: flex; gap: 8px;">
                <button id="selectAllBtn" class="${styles.secondaryButton}">Select All</button>
                <button id="deselectAllBtn" class="${styles.secondaryButton}">Deselect All</button>
              </div>
            </div>
            <div style="flex-shrink: 0;">
              <button id="processButton" class="${styles.primaryButton}">
                Process Selected Files
              </button>
            </div>
          </div>

          <div style="background-color: #f8f7f6; border: 1px solid #edebe9; border-radius: 4px; padding: 12px 16px; margin-bottom: 16px;">
            <label style="display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 14px; color: #323130;">
              <input type="checkbox" id="preserveHiddenRows" checked style="width: 18px; height: 18px; cursor: pointer;" />
              <span style="font-weight: 600;">Preserve Hidden Rows</span>
              <span style="color: #605e5c; font-weight: normal;">- Keep hidden rows hidden in the summary file</span>
            </label>
          </div>

          <table id="fileTable" class="${styles.fileTable}">
            <thead>
              <tr>
                <th style="width: 50px;">#</th>
                <th style="width: 60px;">Select</th>
                <th>File Name</th>
                <th style="width: 100px;">Has Sheet1</th>
                <th style="width: 120px;">Headers Match</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody id="fileTableBody">
              <!-- Files will be populated here -->
            </tbody>
          </table>

          <div class="${styles.infoSection}">
            <div class="${styles.info}">
              <strong>Note:</strong> Selected files will be combined into ${this.SUMMARY_FILE_NAME} in alphabetical order. 
              Existing data will be cleared (header row preserved).
            </div>
          </div>
        </div>
      </div>
    `;

    // Attach event listeners
    this.attachEventListeners();

    // Load files
    await this.loadAndValidateFiles();
  }

  private attachEventListeners(): void {
    const processBtn = this.domElement.querySelector('#processButton');
    const selectAllBtn = this.domElement.querySelector('#selectAllBtn');
    const deselectAllBtn = this.domElement.querySelector('#deselectAllBtn');

    if (processBtn) {
      processBtn.addEventListener('click', () => this.showConfirmDialog());
    }

    if (selectAllBtn) {
      selectAllBtn.addEventListener('click', () => this.toggleAllCheckboxes(true));
    }

    if (deselectAllBtn) {
      deselectAllBtn.addEventListener('click', () => this.toggleAllCheckboxes(false));
    }
  }

  private toggleAllCheckboxes(checked: boolean): void {
    const checkboxes = this.domElement.querySelectorAll<HTMLInputElement>('input[type="checkbox"][data-file-index]');
    checkboxes.forEach(cb => {
      // Only toggle enabled checkboxes (valid files)
      if (!cb.disabled) {
        cb.checked = checked;
      }
    });
  }

  private toggleSpinner(show: boolean, message: string = 'Loading...'): void {
    const spinner = this.domElement.querySelector('#loadingSpinner') as HTMLElement;
    const messageEl = this.domElement.querySelector('#loadingMessage') as HTMLElement;
    
    if (spinner) {
      spinner.style.display = show ? 'block' : 'none';
    }
    if (messageEl) {
      messageEl.textContent = message;
    }
  }

  private showError(message: string): void {
    const errorEl = this.domElement.querySelector('#errorMessage') as HTMLElement;
    if (errorEl) {
      errorEl.textContent = `❌ Error: ${message}`;
      errorEl.style.display = 'block';
    }
  }

  private showWarning(message: string): void {
    const warningEl = this.domElement.querySelector('#warningMessage') as HTMLElement;
    if (warningEl) {
      warningEl.textContent = `⚠️ Warning: ${message}`;
      warningEl.style.display = 'block';
    }
  }

  private hideMessages(): void {
    const errorEl = this.domElement.querySelector('#errorMessage') as HTMLElement;
    const warningEl = this.domElement.querySelector('#warningMessage') as HTMLElement;
    
    if (errorEl) errorEl.style.display = 'none';
    if (warningEl) warningEl.style.display = 'none';
  }

  private updateFileCount(): void {
    const fileCountEl = this.domElement.querySelector('#fileCount') as HTMLElement;
    if (fileCountEl) {
      fileCountEl.textContent = this.excelFiles.length.toString();
    }
  }

  private getPreserveHiddenRowsOption(): boolean {
    const checkbox = this.domElement.querySelector('#preserveHiddenRows') as HTMLInputElement;
    return checkbox ? checkbox.checked : false;
  }

  private showConfirmDialog(): void {
    const selectedFiles = this.getSelectedFiles();

    if (selectedFiles.length === 0) {
      alert('Please select at least one file to process.');
      return;
    }

    const preserveHidden = this.getPreserveHiddenRowsOption();

    const message = 
      `⚠️ CONFIRM DATA PROCESSING\n\n` +
      `You are about to process ${selectedFiles.length} file(s).\n\n` +
      `This action will:\n` +
      `• Clear all existing data from ${this.SUMMARY_FILE_NAME}\n` +
      `• Keep the header row intact\n` +
      `• Add data from the selected files in alphabetical order\n` +
      (preserveHidden ? `• Preserve hidden rows (keep them hidden)\n` : `• Copy hidden rows as visible\n`) +
      `\n` +
      `⚠️ WARNING: Existing data will be permanently deleted.\n` +
      `This action cannot be undone.\n\n` +
      `Do you want to continue?`;

    const confirmed = confirm(message);

    if (confirmed) {
      this.processSelectedFiles().catch((error: Error) => {
        console.error('[RiskCentre] Error in processSelectedFiles:', error);
      });
    }
  }

  private async loadAndValidateFiles(): Promise<void> {
    this.toggleSpinner(true, 'Loading files from Risk per centre folder...');
    this.hideMessages();

    try {
      // Get folder path
      const folderUrl = `${this.context.pageContext.web.serverRelativeUrl}/${this.FOLDER_PATH}`;
      
      console.log('[RiskCentre] Loading files from:', folderUrl);

      // Get folder
      const folder = await this.sp.web.getFolderByServerRelativePath(folderUrl);
      const files = await folder.files();

      console.log('[RiskCentre] Found files:', files.length);

      // Filter Excel files (exclude summary file)
      const excelFiles = files.filter(file => 
        file.Name !== this.SUMMARY_FILE_NAME &&
        (file.Name.endsWith('.xlsx') || file.Name.endsWith('.xls'))
      );

      if (excelFiles.length === 0) {
        this.showWarning('No Excel files found in the folder (excluding _OVERVIEW SUMMARY.xlsx)');
        this.toggleSpinner(false);
        return;
      }

      console.log('[RiskCentre] Excel files to process:', excelFiles.length);

      // Load summary file to get expected headers
      await this.loadExpectedHeaders();

      // Validate each file
      this.excelFiles = [];
      for (const file of excelFiles) {
        const fileInfo = await this.validateFile(file);
        this.excelFiles.push(fileInfo);
      }

      // Sort files alphabetically by name
      this.excelFiles.sort((a, b) => a.name.localeCompare(b.name));

      console.log('[RiskCentre] Files sorted alphabetically');

      // Render table
      this.renderFileTable();

      // Update file count
      this.updateFileCount();

      // Show warnings if any files have issues
      const invalidFiles = this.excelFiles.filter(f => !f.hasSheet1 || !f.headerMatch);
      if (invalidFiles.length > 0) {
        this.showWarning(
          `${invalidFiles.length} file(s) have validation issues. They are disabled and cannot be selected.`
        );
      }

      this.toggleSpinner(false);

    } catch (error) {
      console.error('[RiskCentre] Error loading files:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.showError(`Failed to load files: ${errorMessage}`);
      this.toggleSpinner(false);
    }
  }

  private async loadExpectedHeaders(): Promise<void> {
    try {
      console.log('[RiskCentre] Loading expected headers from:', this.summaryFilePath);

      const fileBuffer = await this.sp.web
        .getFileByServerRelativePath(this.summaryFilePath)
        .getBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(fileBuffer);

      const worksheet = workbook.getWorksheet(this.SHEET_NAME);
      
      if (!worksheet) {
        throw new Error(`Summary file does not contain "${this.SHEET_NAME}" sheet`);
      }

      // Get headers from first row
      const headerRow = worksheet.getRow(1);
      this.expectedHeaders = [];
      
      headerRow.eachCell((cell) => {
        const value = cell.value;
        if (value !== null && value !== undefined) {
          this.expectedHeaders.push(String(value).trim());
        }
      });

      console.log('[RiskCentre] Expected headers:', this.expectedHeaders);

      if (this.expectedHeaders.length === 0) {
        throw new Error('Summary file header row is empty');
      }

    } catch (error) {
      console.error('[RiskCentre] Error loading expected headers:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      throw new Error(`Cannot load summary file: ${errorMessage}`);
    }
  }

  private async validateFile(file: {Name: string; ServerRelativeUrl: string}): Promise<IExcelFileInfo> {
    const fileInfo: IExcelFileInfo = {
      name: file.Name,
      url: file.ServerRelativeUrl,
      hasSheet1: false,
      headerMatch: false
    };

    try {
      // Load file
      const fileBuffer = await this.sp.web
        .getFileByServerRelativePath(file.ServerRelativeUrl)
        .getBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(fileBuffer);

      // Check Sheet1 exists
      const worksheet = workbook.getWorksheet(this.SHEET_NAME);
      if (!worksheet) {
        fileInfo.errorMessage = 'Missing Sheet1';
        return fileInfo;
      }

      fileInfo.hasSheet1 = true;

      // Check headers match
      const headerRow = worksheet.getRow(1);
      const fileHeaders: string[] = [];
      
      headerRow.eachCell((cell) => {
        const value = cell.value;
        if (value !== null && value !== undefined) {
          fileHeaders.push(String(value).trim());
        }
      });

      // Compare headers
      if (fileHeaders.length !== this.expectedHeaders.length) {
        fileInfo.errorMessage = `Header count mismatch (expected ${this.expectedHeaders.length}, got ${fileHeaders.length})`;
        return fileInfo;
      }

      const headersMatch = fileHeaders.every((header, index) => 
        header === this.expectedHeaders[index]
      );

      if (!headersMatch) {
        fileInfo.errorMessage = 'Headers do not match summary file';
        return fileInfo;
      }

      fileInfo.headerMatch = true;

    } catch (error) {
      console.error(`[RiskCentre] Error validating file ${file.Name}:`, error);
      const errorMessage = error instanceof Error ? error.message : 'Validation error';
      fileInfo.errorMessage = errorMessage;
    }

    return fileInfo;
  }

  private renderFileTable(): void {
    const tableBody = this.domElement.querySelector('#fileTableBody') as HTMLTableSectionElement;
    const tableContainer = this.domElement.querySelector('#fileTableContainer') as HTMLElement;

    if (!tableBody || !tableContainer) return;

    tableBody.innerHTML = '';
    
    if (this.excelFiles.length === 0) {
      tableBody.innerHTML = '<tr><td colspan="6" style="text-align: center;">No Excel files found</td></tr>';
      return;
    }

    this.excelFiles.forEach((file, index) => {
      const isValid = file.hasSheet1 && file.headerMatch;
      const row = document.createElement('tr');
      
      if (!isValid) {
        row.classList.add(styles.invalidRow);
      }

      row.innerHTML = `
        <td>${index + 1}</td>
        <td style="text-align: center;">
          <input 
            type="checkbox" 
            data-file-index="${index}" 
            ${isValid ? 'checked' : 'disabled'}
          />
        </td>
        <td>${file.name}</td>
        <td style="text-align: center;">${file.hasSheet1 ? '✅' : '❌'}</td>
        <td style="text-align: center;">${file.headerMatch ? '✅' : '❌'}</td>
        <td>${isValid ? '✅ Ready' : `❌ ${file.errorMessage || 'Invalid'}`}</td>
      `;
      
      tableBody.appendChild(row);
    });

    tableContainer.style.display = 'block';
  }

  private getSelectedFiles(): IExcelFileInfo[] {
    const checkboxes = this.domElement.querySelectorAll<HTMLInputElement>('input[type="checkbox"][data-file-index]:checked');
    const selectedIndexes = Array.from(checkboxes)
      .map(cb => parseInt(cb.getAttribute('data-file-index') || '-1'))
      .filter(index => index >= 0);

    // Get selected files and sort them alphabetically to maintain order
    const selectedFiles = this.excelFiles.filter((_, index) => selectedIndexes.includes(index));
    
    // Sort selected files alphabetically by name
    selectedFiles.sort((a, b) => a.name.localeCompare(b.name));

    return selectedFiles;
  }

  private async processSelectedFiles(): Promise<void> {
    const selectedFiles = this.getSelectedFiles();

    if (selectedFiles.length === 0) {
      alert('Please select at least one file to process.');
      return;
    }

    // Get preserve hidden rows option
    const preserveHiddenRows = this.getPreserveHiddenRowsOption();

    this.toggleSpinner(true, `Processing ${selectedFiles.length} file(s)...`);
    this.hideMessages();

    try {
      console.log('[RiskCentre] Processing files in alphabetical order:', selectedFiles.map(f => f.name));
      console.log('[RiskCentre] Preserve hidden rows:', preserveHiddenRows);

      // Load summary file
      const summaryBuffer = await this.sp.web
        .getFileByServerRelativePath(this.summaryFilePath)
        .getBuffer();

      const summaryWorkbook = new ExcelJS.Workbook();
      await summaryWorkbook.xlsx.load(summaryBuffer);

      const summarySheet = summaryWorkbook.getWorksheet(this.SHEET_NAME);

      if (!summarySheet) {
        throw new Error(`Summary file does not have "${this.SHEET_NAME}" sheet`);
      }

      // Clear existing data (keep header row)
      console.log('[RiskCentre] Clearing existing data...');
      
      const rowCount = summarySheet.rowCount;
      for (let i = rowCount; i > 1; i--) {
        summarySheet.spliceRows(i, 1);
      }

      console.log('[RiskCentre] Data cleared, header preserved');

      // Process each selected file in alphabetical order
      let totalRowsAdded = 0;
      let hiddenRowsCount = 0;

      for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        
        this.toggleSpinner(true, `Processing file ${i + 1}/${selectedFiles.length}: ${file.name}...`);
        
        console.log(`[RiskCentre] Processing [${i + 1}/${selectedFiles.length}]: ${file.name}`);

        // Load source file
        const sourceBuffer = await this.sp.web
          .getFileByServerRelativePath(file.url)
          .getBuffer();

        const sourceWorkbook = new ExcelJS.Workbook();
        await sourceWorkbook.xlsx.load(sourceBuffer);

        const sourceSheet = sourceWorkbook.getWorksheet(this.SHEET_NAME);

        if (!sourceSheet) {
          console.warn(`[RiskCentre] ${file.name} missing Sheet1, skipping`);
          continue;
        }

        // Copy rows (skip header row 1)
        let rowsFromFile = 0;
        sourceSheet.eachRow((sourceRow, rowNumber) => {
          if (rowNumber > 1) { // Skip header
            // Add new row with values
            const newRow = summarySheet.addRow(sourceRow.values);

            // Copy row height
            if (sourceRow.height !== undefined && sourceRow.height !== null) {
              newRow.height = sourceRow.height;
            }

            // Copy row hidden state (only if option is enabled)
            if (preserveHiddenRows && sourceRow.hidden === true) {
              newRow.hidden = true;
              hiddenRowsCount++;
            }

            rowsFromFile++;
            totalRowsAdded++;
          }
        });

        console.log(`[RiskCentre] Added ${rowsFromFile} rows from ${file.name}`);
      }

      // Save summary file
      this.toggleSpinner(true, 'Saving summary file...');
      console.log('[RiskCentre] Saving summary file with', totalRowsAdded, 'total rows in alphabetical order');

      const summaryBufferNew = await summaryWorkbook.xlsx.writeBuffer();

      await this.sp.web
        .getFileByServerRelativePath(this.summaryFilePath)
        .setContent(summaryBufferNew);

      this.toggleSpinner(false);

      let successMessage = 
        `✅ Success!\n\n` +
        `Processed ${selectedFiles.length} file(s) in alphabetical order\n` +
        `Added ${totalRowsAdded} data rows to ${this.SUMMARY_FILE_NAME}`;

      if (preserveHiddenRows && hiddenRowsCount > 0) {
        successMessage += `\n(including ${hiddenRowsCount} hidden rows)`;
      }

      alert(successMessage);

      console.log('[RiskCentre] Processing completed successfully');

    } catch (error) {
      console.error('[RiskCentre] Error processing files:', error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      this.showError(`Failed to process files: ${errorMessage}`);
      this.toggleSpinner(false);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Risk per Centre Aggregator Settings'
          },
          groups: [
            {
              groupName: 'Configuration',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description',
                  value: 'Aggregates Excel files from Risk per centre folder'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}