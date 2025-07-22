import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { SheetData, ParsedFile, ValidationError, ValidationSummary } from '@/types/validation';

export const parseFile = async (file: File): Promise<ParsedFile> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        let sheets: SheetData[] = [];
        
        if (file.name.endsWith('.csv')) {
          // Parse CSV file
          const csvData = Papa.parse(data as string, {
            header: false,
            skipEmptyLines: true,
          });
          
          sheets = [{
            name: 'Sheet1',
            data: csvData.data as string[][],
            headers: csvData.data[0] as string[] || [],
          }];
        } else {
          // Parse Excel file
          const workbook = XLSX.read(data, { type: 'array' });
          
          sheets = workbook.SheetNames.map(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            const arrayData = jsonData as any[][];
            
            return {
              name: sheetName,
              data: arrayData,
              headers: arrayData[0] || [],
            };
          });
        }
        
        // Run validation
        const validationResults = validateFMBDumpSheet(sheets);
        
        const parsedFile: ParsedFile = {
          fileName: file.name,
          sheets,
          isValid: validationResults.summary.totalErrors === 0,
          validationErrors: validationResults.errors,
          validationSummary: validationResults.summary,
        };
        
        resolve(parsedFile);
      } catch (error) {
        reject(new Error(`Failed to parse file: ${error}`));
      }
    };
    
    reader.onerror = () => {
      reject(new Error('Failed to read file'));
    };
    
    if (file.name.endsWith('.csv')) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
};


const validateFMBDumpSheet = (sheets: SheetData[]): { 
  errors: ValidationError[], 
  summary: ValidationSummary 
} => {
  const errors: ValidationError[] = [];
  const errorsBySheet: Record<string, number> = {};
  const warningsBySheet: Record<string, number> = {};

  // Expected sheet names for FMB dump (exact names required)
  const expectedSheets = [
    'Designation Mapping',
    'Survey Master',
    'Question Master',
    'Access Sheet'
  ];

  // Check if we have exactly 4 sheets with correct names
  if (sheets.length !== 4) {
    const error: ValidationError = {
      id: `incorrect-sheet-count`,
      sheet: 'File Structure',
      row: 0,
      column: '',
      field: 'Sheet Count',
      message: `File must contain exactly 4 sheets. Found ${sheets.length} sheets.`,
      severity: 'error',
      rule: 'SHEET_COUNT_001',
    };
    errors.push(error);
  }

  // Check if required sheets exist with exact names
  expectedSheets.forEach(expectedSheet => {
    const found = sheets.find(sheet => 
      sheet.name.trim() === expectedSheet
    );
    
    if (!found) {
      const error: ValidationError = {
        id: `missing-sheet-${expectedSheet}`,
        sheet: 'File Structure',
        row: 0,
        column: '',
        field: 'Required Sheet',
        message: `Missing required sheet: "${expectedSheet}". Sheet names must be exact.`,
        severity: 'error',
        rule: 'SHEET_STRUCTURE_001',
      };
      errors.push(error);
    }
  });

  // Check for unexpected sheets
  sheets.forEach(sheet => {
    if (!expectedSheets.includes(sheet.name.trim())) {
      const error: ValidationError = {
        id: `unexpected-sheet-${sheet.name}`,
        sheet: 'File Structure',
        row: 0,
        column: '',
        field: 'Unexpected Sheet',
        message: `Unexpected sheet: "${sheet.name}". Only these sheets are allowed: ${expectedSheets.join(', ')}`,
        severity: 'error',
        rule: 'SHEET_STRUCTURE_002',
      };
      errors.push(error);
    }
  });

  // Validate each sheet
  sheets.forEach(sheet => {
    const sheetErrors = validateSheet(sheet);
    errors.push(...sheetErrors);
  });

  // Count errors and warnings by sheet
  errors.forEach(error => {
    if (error.severity === 'error') {
      errorsBySheet[error.sheet] = (errorsBySheet[error.sheet] || 0) + 1;
    } else if (error.severity === 'warning') {
      warningsBySheet[error.sheet] = (warningsBySheet[error.sheet] || 0) + 1;
    }
  });

  const summary: ValidationSummary = {
    totalErrors: errors.filter(e => e.severity === 'error').length,
    totalWarnings: errors.filter(e => e.severity === 'warning').length,
    errorsBySheet,
    warningsBySheet,
  };

  return { errors, summary };
};

const validateSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  
  // Basic validation - check for empty rows and required fields
  if (sheet.data.length === 0) {
    errors.push({
      id: `empty-sheet-${sheet.name}`,
      sheet: sheet.name,
      row: 0,
      column: '',
      field: 'Sheet Data',
      message: 'Sheet is empty',
      severity: 'error',
      rule: 'SHEET_EMPTY_001',
    });
    return errors;
  }

  // Check for headers
  if (sheet.headers.length === 0) {
    errors.push({
      id: `no-headers-${sheet.name}`,
      sheet: sheet.name,
      row: 1,
      column: '',
      field: 'Headers',
      message: 'No headers found in sheet',
      severity: 'error',
      rule: 'HEADERS_001',
    });
  }

  // Sheet-specific validations (exact name matching)
  if (sheet.name.trim() === 'Designation Mapping') {
    return validateDesignationMappingSheet(sheet);
  } else if (sheet.name.trim() === 'Survey Master') {
    return validateSurveyMasterSheet(sheet);
  } else if (sheet.name.trim() === 'Question Master') {
    return validateQuestionMasterSheet(sheet);
  } else if (sheet.name.trim() === 'Access Sheet') {
    return validateAccessSheet(sheet);
  }

  return errors;
};

const validateDesignationMappingSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  
  // Check required columns - exactly 5 columns
  const requiredColumns = ['State', 'Medium', 'Medium_in_english', 'List of Designations', 'Hierarchical Level'];
  
  if (sheet.headers.length !== 5) {
    errors.push({
      id: `DM-COLCOUNT-${Date.now()}`,
      sheet: sheet.name,
      row: 1,
      column: 'Headers',
      field: 'Column Count',
      message: `Expected exactly 5 columns, found ${sheet.headers.length}`,
      severity: 'error',
      rule: 'DM-00'
    });
  }

  const missingColumns = requiredColumns.filter(col => !sheet.headers.includes(col));
  if (missingColumns.length > 0) {
    errors.push({
      id: `DM-COLS-${Date.now()}`,
      sheet: sheet.name,
      row: 1,
      column: 'Headers',
      field: 'Required Columns',
      message: `Missing required columns: ${missingColumns.join(', ')}`,
      severity: 'error',
      rule: 'DM-01'
    });
  }

  // Track hierarchical levels for duplicate validation
  const hierarchicalLevels: { [key: string]: { level: number, medium: string, count: number }[] } = {};

  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNum = index + 2;
    const stateCol = sheet.headers.indexOf('State');
    const mediumCol = sheet.headers.indexOf('Medium');
    const mediumEngCol = sheet.headers.indexOf('Medium_in_english');
    const designationsCol = sheet.headers.indexOf('List of Designations');
    const hierarchicalCol = sheet.headers.indexOf('Hierarchical Level');

    // Check for trailing whitespaces in all cells
    row.forEach((cell, cellIndex) => {
      if (typeof cell === 'string' && (cell.startsWith(' ') || cell.endsWith(' '))) {
        errors.push({
          id: `DM-WS-${rowNum}-${cellIndex}`,
          sheet: sheet.name,
          row: rowNum,
          column: sheet.headers[cellIndex],
          field: sheet.headers[cellIndex],
          message: 'Cell contains leading or trailing whitespace',
          severity: 'error',
          rule: 'U-01'
        });
      }
    });

    // State validation - Indian states in English only
    const state = row[stateCol];
    if (!state) {
      errors.push({
        id: `DM-STATE-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'State',
        field: 'State',
        message: 'State is mandatory',
        severity: 'error',
        rule: 'DM-02'
      });
    } else if (typeof state !== 'string' || !/^[A-Za-z\s]+$/.test(state)) {
      errors.push({
        id: `DM-STATE-FORMAT-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'State',
        field: 'State',
        message: 'State must contain only English letters and spaces',
        severity: 'error',
        rule: 'DM-03'
      });
    }

    // Medium validation
    const medium = row[mediumCol];
    if (!medium) {
      errors.push({
        id: `DM-MEDIUM-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'Medium',
        field: 'Medium',
        message: 'Medium is mandatory',
        severity: 'error',
        rule: 'DM-04'
      });
    }

    // Medium_in_english validation
    const mediumEng = row[mediumEngCol];
    if (!mediumEng) {
      errors.push({
        id: `DM-MEDIUMENG-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'Medium_in_english',
        field: 'Medium_in_english',
        message: 'Medium_in_english is mandatory',
        severity: 'error',
        rule: 'DM-05'
      });
    }

    // List of Designations validation - alphanumeric with only '_' special character
    const designations = row[designationsCol];
    if (!designations) {
      errors.push({
        id: `DM-DESIGNATIONS-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'List of Designations',
        field: 'List of Designations',
        message: 'List of Designations is mandatory',
        severity: 'error',
        rule: 'DM-06'
      });
    } else if (typeof designations === 'string' && !/^[A-Za-z0-9_\u0900-\u097F\u0A80-\u0AFF\u0980-\u09FF]+$/.test(designations)) {
      errors.push({
        id: `DM-DESIGNATIONS-FORMAT-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'List of Designations',
        field: 'List of Designations',
        message: 'List of Designations must be alphanumeric with only underscore (_) as special character',
        severity: 'error',
        rule: 'DM-07'
      });
    }

    // Hierarchical Level validation
    const hierarchicalLevel = row[hierarchicalCol];
    if (hierarchicalLevel === undefined || hierarchicalLevel === null || hierarchicalLevel === '') {
      errors.push({
        id: `DM-HIERARCHY-${rowNum}`,
        sheet: sheet.name,
        row: rowNum,
        column: 'Hierarchical Level',
        field: 'Hierarchical Level',
        message: 'Hierarchical Level is mandatory',
        severity: 'error',
        rule: 'DM-08'
      });
    } else {
      const level = Number(hierarchicalLevel);
      if (isNaN(level) || level < 0 || !Number.isInteger(level)) {
        errors.push({
          id: `DM-HIERARCHY-FORMAT-${rowNum}`,
          sheet: sheet.name,
          row: rowNum,
          column: 'Hierarchical Level',
          field: 'Hierarchical Level',
          message: 'Hierarchical Level must be a numeric value from 0 to positive numbers only',
          severity: 'error',
          rule: 'DM-09'
        });
      } else {
        // Track for duplicate validation
        const key = `${level}`;
        if (!hierarchicalLevels[key]) {
          hierarchicalLevels[key] = [];
        }
        hierarchicalLevels[key].push({ level, medium: medium as string, count: 1 });
      }
    }
  });

  // Check for hierarchical level duplicates
  Object.keys(hierarchicalLevels).forEach(levelKey => {
    const entries = hierarchicalLevels[levelKey];
    if (entries.length > 2) {
      errors.push({
        id: `DM-HIERARCHY-DUP-${levelKey}`,
        sheet: sheet.name,
        row: 0,
        column: 'Hierarchical Level',
        field: 'Hierarchical Level',
        message: `Hierarchical Level ${levelKey} appears more than twice`,
        severity: 'error',
        rule: 'DM-10'
      });
    } else if (entries.length === 2) {
      // Check if both have same medium
      if (entries[0].medium === entries[1].medium) {
        errors.push({
          id: `DM-HIERARCHY-SAME-MEDIUM-${levelKey}`,
          sheet: sheet.name,
          row: 0,
          column: 'Hierarchical Level',
          field: 'Hierarchical Level',
          message: `Hierarchical Level ${levelKey} has duplicate entries with same medium`,
          severity: 'error',
          rule: 'DM-11'
        });
      }
    }
  });

  return errors;
};

const validateSurveyMasterSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['Survey ID', 'Survey Name', 'Survey Description', 'Available_mediums', 'Hierarchical Access Level', 'In School', 'Accept multiple Entries', 'Launch Date', 'Close Date', 'Mode', 'Visible_on_report_bot', 'Is Active?', 'Download_response', 'Geo Fencing', 'Geo Tagging', 'Test Survey'];
  
  // Check required columns
  requiredColumns.forEach(col => {
    if (!sheet.headers.includes(col)) {
      errors.push({
        id: `missing-col-${col}`,
        sheet: sheet.name,
        row: 1,
        column: col,
        field: 'Required Column',
        message: `Missing required column: ${col}`,
        severity: 'error',
        rule: 'SURVEY_MASTER_001',
      });
    }
  });

  const surveyIds = new Set();
  
  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;

    // Check for trailing whitespaces in all cells
    row.forEach((cell, cellIndex) => {
      if (typeof cell === 'string' && (cell.startsWith(' ') || cell.endsWith(' '))) {
        errors.push({
          id: `SM-WS-${rowNumber}-${cellIndex}`,
          sheet: sheet.name,
          row: rowNumber,
          column: sheet.headers[cellIndex],
          field: sheet.headers[cellIndex],
          message: 'Cell contains leading or trailing whitespace',
          severity: 'error',
          rule: 'U-01'
        });
      }
    });
    
    // Survey ID validation - alphanumeric with '_', no spaces
    const surveyIdIndex = sheet.headers.indexOf('Survey ID');
    if (surveyIdIndex >= 0) {
      const surveyId = row[surveyIdIndex]?.toString().trim() || '';
      if (!surveyId) {
        errors.push({
          id: `SM-SURVEYID-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Survey ID',
          field: 'Survey ID',
          message: 'Survey ID is mandatory',
          severity: 'error',
          rule: 'SM-01'
        });
      } else {
        // Check uniqueness
        if (surveyIds.has(surveyId)) {
          errors.push({
            id: `duplicate-survey-id-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Survey ID',
            field: 'Survey ID',
            message: `Duplicate Survey ID: ${surveyId}`,
            severity: 'error',
            rule: 'SM-02',
          });
        } else {
          surveyIds.add(surveyId);
        }
        
        // Validate format - alphanumeric with '_', no spaces
        if (!/^[A-Za-z0-9_]+$/.test(surveyId) || /\s/.test(surveyId)) {
          errors.push({
            id: `invalid-survey-id-format-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Survey ID',
            field: 'Survey ID',
            message: 'Survey ID must be alphanumeric with underscore only, no spaces',
            severity: 'error',
            rule: 'SM-03',
          });
        }
      }
    }

    // Survey Name validation - alphanumeric as single sentence
    const surveyNameIndex = sheet.headers.indexOf('Survey Name');
    if (surveyNameIndex >= 0) {
      const surveyName = row[surveyNameIndex]?.toString().trim() || '';
      if (!surveyName) {
        errors.push({
          id: `SM-NAME-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Survey Name',
          field: 'Survey Name',
          message: 'Survey Name is mandatory',
          severity: 'error',
          rule: 'SM-04'
        });
      }
    }

    // Survey Description validation - up to 60 words
    const surveyDescIndex = sheet.headers.indexOf('Survey Description');
    if (surveyDescIndex >= 0) {
      const surveyDesc = row[surveyDescIndex]?.toString().trim() || '';
      if (surveyDesc) {
        const wordCount = surveyDesc.split(/\s+/).length;
        if (wordCount > 60) {
          errors.push({
            id: `SM-DESC-WORDS-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Survey Description',
            field: 'Survey Description',
            message: `Survey Description exceeds 60 words limit (${wordCount} words)`,
            severity: 'error',
            rule: 'SM-05'
          });
        }
      }
    }

    // Available_mediums validation - comma separated, no spaces
    const availableMediumsIndex = sheet.headers.indexOf('Available_mediums');
    if (availableMediumsIndex >= 0) {
      const availableMediums = row[availableMediumsIndex]?.toString().trim() || '';
      if (availableMediums && /\s/.test(availableMediums)) {
        errors.push({
          id: `SM-MEDIUMS-SPACES-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Available_mediums',
          field: 'Available_mediums',
          message: 'Available_mediums must be comma separated with no spaces',
          severity: 'error',
          rule: 'SM-06'
        });
      }
    }

    // Hierarchical Access Level validation - comma separated, no spaces
    const hierarchicalAccessIndex = sheet.headers.indexOf('Hierarchical Access Level');
    if (hierarchicalAccessIndex >= 0) {
      const hierarchicalAccess = row[hierarchicalAccessIndex]?.toString().trim() || '';
      if (hierarchicalAccess && /\s/.test(hierarchicalAccess)) {
        errors.push({
          id: `SM-HIERARCHY-SPACES-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Hierarchical Access Level',
          field: 'Hierarchical Access Level',
          message: 'Hierarchical Access Level must be comma separated with no spaces',
          severity: 'error',
          rule: 'SM-07'
        });
      }
    }

    // Launch Date validation - DD/MM/YYYY 00:00:00 format
    const launchDateIndex = sheet.headers.indexOf('Launch Date');
    if (launchDateIndex >= 0) {
      const launchDate = row[launchDateIndex]?.toString().trim() || '';
      if (!launchDate) {
        errors.push({
          id: `SM-LAUNCH-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Launch Date',
          field: 'Launch Date',
          message: 'Launch Date is mandatory',
          severity: 'error',
          rule: 'SM-08'
        });
      } else if (!/^\d{2}\/\d{2}\/\d{4} 00:00:00$/.test(launchDate)) {
        errors.push({
          id: `SM-LAUNCH-FORMAT-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Launch Date',
          field: 'Launch Date',
          message: 'Launch Date must be in DD/MM/YYYY 00:00:00 format',
          severity: 'error',
          rule: 'SM-09'
        });
      }
    }

    // Close Date validation - DD/MM/YYYY 23:59:00 format, can't be in past
    const closeDateIndex = sheet.headers.indexOf('Close Date');
    if (closeDateIndex >= 0) {
      const closeDate = row[closeDateIndex]?.toString().trim() || '';
      if (closeDate && !/^\d{2}\/\d{2}\/\d{4} 23:59:00$/.test(closeDate)) {
        errors.push({
          id: `SM-CLOSE-FORMAT-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Close Date',
          field: 'Close Date',
          message: 'Close Date must be in DD/MM/YYYY 23:59:00 format',
          severity: 'error',
          rule: 'SM-10'
        });
      }
    }

    // Mode validation
    const modeIndex = sheet.headers.indexOf('Mode');
    if (modeIndex >= 0) {
      const mode = row[modeIndex]?.toString().trim() || '';
      const validModes = ['New Data', 'Correction', 'delete data', 'None'];
      if (mode && !validModes.includes(mode)) {
        errors.push({
          id: `SM-MODE-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Mode',
          field: 'Mode',
          message: 'Mode must be one of: New Data, Correction, delete data, None',
          severity: 'error',
          rule: 'SM-11'
        });
      }
    }

    // Yes/No fields validation
    const yesNoFields = ['In School', 'Accept multiple Entries', 'Visible_on_report_bot', 'Is Active?', 'Download_response', 'Geo Tagging'];
    yesNoFields.forEach(field => {
      const fieldIndex = sheet.headers.indexOf(field);
      if (fieldIndex >= 0) {
        const value = row[fieldIndex]?.toString().trim() || '';
        if (value && !['Yes', 'No'].includes(value)) {
          errors.push({
            id: `invalid-yes-no-${field}-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: field,
            field: field,
            message: `${field} must be exactly 'Yes' or 'No'`,
            severity: 'error',
            rule: 'SM-12',
          });
        }
      }
    });

    // Geo Fencing validation - must be 'No' only
    const geoFencingIndex = sheet.headers.indexOf('Geo Fencing');
    if (geoFencingIndex >= 0) {
      const geoFencing = row[geoFencingIndex]?.toString().trim() || '';
      if (geoFencing && geoFencing !== 'No') {
        errors.push({
          id: `SM-GEOFENCING-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Geo Fencing',
          field: 'Geo Fencing',
          message: 'Geo Fencing must be "No" only',
          severity: 'error',
          rule: 'SM-13'
        });
      }
    }

    // Test Survey validation - must be 'No' only
    const testSurveyIndex = sheet.headers.indexOf('Test Survey');
    if (testSurveyIndex >= 0) {
      const testSurvey = row[testSurveyIndex]?.toString().trim() || '';
      if (testSurvey && testSurvey !== 'No') {
        errors.push({
          id: `SM-TESTSURVEY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Test Survey',
          field: 'Test Survey',
          message: 'Test Survey must be "No" only',
          severity: 'error',
          rule: 'SM-14'
        });
      }
    }
  });

  return errors;
};

const validateQuestionMasterSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  
  const questionIds = new Set();
  
  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;

    // Check for trailing whitespaces in all cells
    row.forEach((cell, cellIndex) => {
      if (typeof cell === 'string' && (cell.startsWith(' ') || cell.endsWith(' '))) {
        errors.push({
          id: `QM-WS-${rowNumber}-${cellIndex}`,
          sheet: sheet.name,
          row: rowNumber,
          column: sheet.headers[cellIndex],
          field: sheet.headers[cellIndex],
          message: 'Cell contains leading or trailing whitespace',
          severity: 'error',
          rule: 'U-01'
        });
      }
    });
    
    // Survey ID validation - must match Survey Master
    const surveyIdIndex = sheet.headers.indexOf('Survey ID');
    if (surveyIdIndex >= 0) {
      const surveyId = row[surveyIdIndex]?.toString().trim() || '';
      if (!surveyId) {
        errors.push({
          id: `QM-SURVEYID-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Survey ID',
          field: 'Survey ID',
          message: 'Survey ID is mandatory',
          severity: 'error',
          rule: 'QM-01'
        });
      }
    }

    // Medium validation
    const mediumIndex = sheet.headers.indexOf('Medium');
    if (mediumIndex >= 0) {
      const medium = row[mediumIndex]?.toString().trim() || '';
      if (!medium) {
        errors.push({
          id: `QM-MEDIUM-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Medium',
          field: 'Medium',
          message: 'Medium is mandatory',
          severity: 'error',
          rule: 'QM-02'
        });
      }
    }

    // Medium_in_english validation
    const mediumEngIndex = sheet.headers.indexOf('Medium_in_english');
    if (mediumEngIndex >= 0) {
      const mediumEng = row[mediumEngIndex]?.toString().trim() || '';
      if (!mediumEng) {
        errors.push({
          id: `QM-MEDIUMENG-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Medium_in_english',
          field: 'Medium_in_english',
          message: 'Medium_in_english is mandatory',
          severity: 'error',
          rule: 'QM-03'
        });
      }
    }
    
    // Question_ID validation - Q1, Q2, Q1.1, Q2.a, Q2.1.a format
    const questionIdIndex = sheet.headers.indexOf('Question_ID');
    if (questionIdIndex >= 0) {
      const questionId = row[questionIdIndex]?.toString().trim() || '';
      if (!questionId) {
        errors.push({
          id: `QM-QUESTIONID-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_ID',
          field: 'Question_ID',
          message: 'Question_ID is mandatory and cannot be empty',
          severity: 'error',
          rule: 'QM-04'
        });
      } else {
        // Check uniqueness
        if (questionIds.has(questionId)) {
          errors.push({
            id: `duplicate-question-id-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Question_ID',
            field: 'Question_ID',
            message: `Duplicate Question ID: ${questionId}`,
            severity: 'error',
            rule: 'QM-05',
          });
        } else {
          questionIds.add(questionId);
        }
        
        // Validate Question_ID format - Q followed by number, optionally with sub-questions
        if (!/^Q\d+(\.\d+|\.a|\.1\.a)?$/.test(questionId)) {
          errors.push({
            id: `invalid-question-id-format-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Question_ID',
            field: 'Question_ID',
            message: 'Question ID must follow format Q1, Q2, Q1.1, Q2.a, Q2.1.a',
            severity: 'error',
            rule: 'QM-06',
          });
        }
      }
    }

    // Question Type validation - predefined types only
    const questionTypeIndex = sheet.headers.indexOf('Question Type');
    if (questionTypeIndex >= 0) {
      const questionType = row[questionTypeIndex]?.toString().trim() || '';
      const validQuestionTypes = [
        'Multiple Choice Single Select',
        'Multiple Choice Multi Select', 
        'Drop Down',
        'Text Response',
        'Tabular Text Input',
        'Tabular Drop Down',
        'Tabular Check Box',
        'Image Upload',
        'Video Upload',
        'Likert Scale',
        'Calendar'
      ];
      if (!questionType) {
        errors.push({
          id: `QM-QUESTIONTYPE-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question Type',
          field: 'Question Type',
          message: 'Question Type is mandatory',
          severity: 'error',
          rule: 'QM-07'
        });
      } else if (!validQuestionTypes.includes(questionType)) {
        errors.push({
          id: `QM-QUESTIONTYPE-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question Type',
          field: 'Question Type',
          message: `Question Type must be one of the predefined types: ${validQuestionTypes.join(', ')}`,
          severity: 'error',
          rule: 'QM-08'
        });
      }
    }

    // IsDynamic validation - must be empty
    const isDynamicIndex = sheet.headers.indexOf('IsDynamic');
    if (isDynamicIndex >= 0) {
      const isDynamic = row[isDynamicIndex]?.toString().trim() || '';
      if (isDynamic) {
        errors.push({
          id: `QM-ISDYNAMIC-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'IsDynamic',
          field: 'IsDynamic',
          message: 'IsDynamic must be empty',
          severity: 'error',
          rule: 'QM-09'
        });
      }
    }

    // Question_Description_Optional validation - max 256 characters
    const questionDescIndex = sheet.headers.indexOf('Question_Description_Optional');
    if (questionDescIndex >= 0) {
      const questionDesc = row[questionDescIndex]?.toString().trim() || '';
      if (questionDesc && questionDesc.length > 256) {
        errors.push({
          id: `QM-QUESTIONDESC-LENGTH-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_Description_Optional',
          field: 'Question_Description_Optional',
          message: 'Question Description Optional exceeds 256 character limit',
          severity: 'error',
          rule: 'QM-10'
        });
      }
    }

    // Max_Value and Min_Value validation
    const maxValueIndex = sheet.headers.indexOf('Max_Value');
    const minValueIndex = sheet.headers.indexOf('Min_Value');
    if (maxValueIndex >= 0 && minValueIndex >= 0) {
      const maxValue = row[maxValueIndex]?.toString().trim() || '';
      const minValue = row[minValueIndex]?.toString().trim() || '';
      
      if (maxValue && minValue) {
        const maxNum = Number(maxValue);
        const minNum = Number(minValue);
        if (!isNaN(maxNum) && !isNaN(minNum) && maxNum <= minNum) {
          errors.push({
            id: `QM-MAXMIN-INVALID-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Max_Value',
            field: 'Max_Value',
            message: 'Max_Value must be greater than Min_Value',
            severity: 'error',
            rule: 'QM-11'
          });
        }
      }
    }

    // Is Mandatory validation - Yes/No only
    const isMandatoryIndex = sheet.headers.indexOf('Is Mandatory');
    if (isMandatoryIndex >= 0) {
      const isMandatory = row[isMandatoryIndex]?.toString().trim() || '';
      if (isMandatory && !['Yes', 'No'].includes(isMandatory)) {
        errors.push({
          id: `QM-ISMANDATORY-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Is Mandatory',
          field: 'Is Mandatory',
          message: 'Is Mandatory must be exactly "Yes" or "No"',
          severity: 'error',
          rule: 'QM-12'
        });
      }
    }

    // Table_Header_value validation - mandatory for tabular types
    const tableHeaderIndex = sheet.headers.indexOf('Table_Header_value');
    const questionTypeIndex2 = sheet.headers.indexOf('Question Type');
    if (tableHeaderIndex >= 0 && questionTypeIndex2 >= 0) {
      const tableHeader = row[tableHeaderIndex]?.toString().trim() || '';
      const questionType = row[questionTypeIndex2]?.toString().trim() || '';
      const tabularTypes = ['Tabular Text Input', 'Tabular Drop Down', 'Tabular Check Box'];
      
      if (tabularTypes.includes(questionType) && !tableHeader) {
        errors.push({
          id: `QM-TABLEHEADER-MISSING-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Table_Header_value',
          field: 'Table_Header_value',
          message: 'Table_Header_value is mandatory for Tabular question types',
          severity: 'error',
          rule: 'QM-13'
        });
      } else if (!tabularTypes.includes(questionType) && tableHeader) {
        errors.push({
          id: `QM-TABLEHEADER-UNEXPECTED-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Table_Header_value',
          field: 'Table_Header_value',
          message: 'Table_Header_value must be blank for non-tabular question types',
          severity: 'error',
          rule: 'QM-14'
        });
      }
    }

    // Mode validation
    const modeIndex = sheet.headers.indexOf('Mode');
    if (modeIndex >= 0) {
      const mode = row[modeIndex]?.toString().trim() || '';
      const validModes = ['New Data', 'Correction', 'Delete Data', 'None'];
      if (mode && !validModes.includes(mode)) {
        errors.push({
          id: `QM-MODE-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Mode',
          field: 'Mode',
          message: 'Mode must be one of: New Data, Correction, Delete Data, None',
          severity: 'error',
          rule: 'QM-15'
        });
      }
    }

    // Question_Media_Type validation - must be "None", not empty
    const mediaTypeIndex = sheet.headers.indexOf('Question_Media_Type');
    if (mediaTypeIndex >= 0) {
      const mediaType = row[mediaTypeIndex]?.toString().trim() || '';
      if (!mediaType || mediaType === '') {
        errors.push({
          id: `QM-MEDIATYPE-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_Media_Type',
          field: 'Question_Media_Type',
          message: 'Question_Media_Type cannot be empty. Use "None" if no media',
          severity: 'error',
          rule: 'QM-16'
        });
      } else if (mediaType !== 'None') {
        errors.push({
          id: `QM-MEDIATYPE-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_Media_Type',
          field: 'Question_Media_Type',
          message: 'Question_Media_Type must be "None" for all questions',
          severity: 'error',
          rule: 'QM-17'
        });
      }
    }

    // Correct_Answer_Optional and Children Questions validation - must be empty
    const correctAnswerIndex = sheet.headers.indexOf('Correct_Answer_Optional');
    const childrenQuestionsIndex = sheet.headers.indexOf('Children Questions');
    
    if (correctAnswerIndex >= 0) {
      const correctAnswer = row[correctAnswerIndex]?.toString().trim() || '';
      if (correctAnswer) {
        errors.push({
          id: `QM-CORRECTANSWER-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Correct_Answer_Optional',
          field: 'Correct_Answer_Optional',
          message: 'Correct_Answer_Optional must be empty',
          severity: 'error',
          rule: 'QM-18'
        });
      }
    }

    if (childrenQuestionsIndex >= 0) {
      const childrenQuestions = row[childrenQuestionsIndex]?.toString().trim() || '';
      if (childrenQuestions) {
        errors.push({
          id: `QM-CHILDRENQUESTIONS-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Children Questions',
          field: 'Children Questions',
          message: 'Children Questions must be empty',
          severity: 'error',
          rule: 'QM-19'
        });
      }
    }
  });

  return errors;
};

const validateAccessSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['Designation', 'Hierarchical Level', 'Bot ID', 'State', 'Name', 'User ID', 'Status'];
  
  // Check required columns
  requiredColumns.forEach(col => {
    if (!sheet.headers.includes(col)) {
      errors.push({
        id: `missing-col-${col}`,
        sheet: sheet.name,
        row: 1,
        column: col,
        field: 'Required Column',
        message: `Missing required column: ${col}`,
        severity: 'error',
        rule: 'ACCESS_SHEET_001',
      });
    }
  });

  const userIds = new Set();
  const stateMobilePairs = new Set();

  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;

    // Check for trailing whitespaces in all cells
    row.forEach((cell, cellIndex) => {
      if (typeof cell === 'string' && (cell.startsWith(' ') || cell.endsWith(' '))) {
        errors.push({
          id: `AS-WS-${rowNumber}-${cellIndex}`,
          sheet: sheet.name,
          row: rowNumber,
          column: sheet.headers[cellIndex],
          field: sheet.headers[cellIndex],
          message: 'Cell contains leading or trailing whitespace',
          severity: 'error',
          rule: 'U-01'
        });
      }
    });
    
    // Designation validation - must match from Designation Mapping
    const designationIndex = sheet.headers.indexOf('Designation');
    if (designationIndex >= 0) {
      const designation = row[designationIndex]?.toString().trim() || '';
      if (!designation) {
        errors.push({
          id: `AS-DESIGNATION-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Designation',
          field: 'Designation',
          message: 'Designation is mandatory',
          severity: 'error',
          rule: 'AS-01'
        });
      }
    }

    // Hierarchical Level validation - must be numeric and match designation mapping
    const hierarchicalLevelIndex = sheet.headers.indexOf('Hierarchical Level');
    if (hierarchicalLevelIndex >= 0) {
      const hierarchicalLevel = row[hierarchicalLevelIndex]?.toString().trim() || '';
      if (!hierarchicalLevel) {
        errors.push({
          id: `AS-HIERARCHY-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Hierarchical Level',
          field: 'Hierarchical Level',
          message: 'Hierarchical Level is mandatory',
          severity: 'error',
          rule: 'AS-02'
        });
      } else {
        const level = Number(hierarchicalLevel);
        if (isNaN(level) || !Number.isInteger(level)) {
          errors.push({
            id: `AS-HIERARCHY-FORMAT-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Hierarchical Level',
            field: 'Hierarchical Level',
            message: 'Hierarchical Level must be a numeric value, not text or special characters',
            severity: 'error',
            rule: 'AS-03'
          });
        }
      }
    }

    // Bot ID validation - optional field, show attention if filled
    const botIdIndex = sheet.headers.indexOf('Bot ID');
    if (botIdIndex >= 0) {
      const botId = row[botIdIndex]?.toString().trim() || '';
      if (botId) {
        errors.push({
          id: `AS-BOTID-ATTENTION-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Bot ID',
          field: 'Bot ID',
          message: 'Bot ID is filled - attention needed',
          severity: 'warning',
          rule: 'AS-04'
        });
      }
    }

    // State validation - must match Designation Mapping
    const stateIndex = sheet.headers.indexOf('State');
    if (stateIndex >= 0) {
      const state = row[stateIndex]?.toString().trim() || '';
      if (!state) {
        errors.push({
          id: `AS-STATE-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'State',
          field: 'State',
          message: 'State is mandatory',
          severity: 'error',
          rule: 'AS-05'
        });
      }
    }

    // Name validation - alpha, numeric, alphanumeric with '_'
    const nameIndex = sheet.headers.indexOf('Name');
    if (nameIndex >= 0) {
      const name = row[nameIndex]?.toString().trim() || '';
      if (!name) {
        errors.push({
          id: `AS-NAME-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Name',
          field: 'Name',
          message: 'Name is mandatory',
          severity: 'error',
          rule: 'AS-06'
        });
      } else if (!/^[A-Za-z0-9_\u0900-\u097F\u0A80-\u0AFF\u0980-\u09FF]+$/.test(name)) {
        errors.push({
          id: `AS-NAME-FORMAT-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Name',
          field: 'Name',
          message: 'Name must be alpha, numeric or alphanumeric with underscore (_) only',
          severity: 'error',
          rule: 'AS-07'
        });
      }
    }

    // User ID validation - unique across dataset, special characters allowed
    const userIdIndex = sheet.headers.indexOf('User ID');
    if (userIdIndex >= 0) {
      const userId = row[userIdIndex]?.toString().trim() || '';
      if (!userId) {
        errors.push({
          id: `AS-USERID-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'User ID',
          field: 'User ID',
          message: 'User ID is mandatory',
          severity: 'error',
          rule: 'AS-08'
        });
      } else {
        // Check uniqueness
        if (userIds.has(userId)) {
          errors.push({
            id: `AS-USERID-DUPLICATE-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'User ID',
            field: 'User ID',
            message: `Duplicate User ID: ${userId}. Each User ID must be unique across the entire dataset`,
            severity: 'error',
            rule: 'AS-09'
          });
        } else {
          userIds.add(userId);
        }
      }
    }

    // Status validation - only "Active" or "Inactive"
    const statusIndex = sheet.headers.indexOf('Status');
    if (statusIndex >= 0) {
      const status = row[statusIndex]?.toString().trim() || '';
      if (!status) {
        errors.push({
          id: `AS-STATUS-EMPTY-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Status',
          field: 'Status',
          message: 'Status is mandatory',
          severity: 'error',
          rule: 'AS-10'
        });
      } else if (!['Active', 'Inactive'].includes(status)) {
        errors.push({
          id: `AS-STATUS-INVALID-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Status',
          field: 'Status',
          message: 'Status must be exactly "Active" or "Inactive"',
          severity: 'error',
          rule: 'AS-11'
        });
      }
    }

    // Validate mobile number format if Mobile column exists
    const mobileIndex = sheet.headers.indexOf('Mobile');
    if (mobileIndex >= 0 && row[mobileIndex]) {
      const mobile = row[mobileIndex].toString().trim();
      const state = row[stateIndex]?.toString().trim() || '';
      
      if (!/^[6-9]\d{9}$/.test(mobile)) {
        errors.push({
          id: `AS-MOBILE-FORMAT-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Mobile',
          field: 'Mobile',
          message: 'Mobile number must be 10 digits starting with 6-9',
          severity: 'error',
          rule: 'AS-12',
        });
      } else {
        // Check (State, Mobile) uniqueness
        const stateMobilePair = `${state}-${mobile}`;
        if (stateMobilePairs.has(stateMobilePair)) {
          errors.push({
            id: `AS-STATEMOBILE-DUPLICATE-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Mobile',
            field: 'Mobile',
            message: `Duplicate (State, Mobile) combination: ${state}, ${mobile}`,
            severity: 'error',
            rule: 'AS-13'
          });
        } else {
          stateMobilePairs.add(stateMobilePair);
        }
      }
    }

    // Validate email format if present
    const emailIndex = sheet.headers.indexOf('Email');
    if (emailIndex >= 0 && row[emailIndex]) {
      const email = row[emailIndex].toString().trim();
      if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        errors.push({
          id: `AS-EMAIL-FORMAT-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Email',
          field: 'Email',
          message: 'Invalid email format',
          severity: 'warning',
          rule: 'AS-14',
        });
      }
    }
  });

  return errors;
};