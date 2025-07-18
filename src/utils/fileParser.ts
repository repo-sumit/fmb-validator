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

  // Expected sheet names for FMB dump
  const expectedSheets = [
    'Designation Mapping Sheet',
    'Survey Master',
    'Question Master',
    'Access Sheet',
    'Question Types'
  ];

  // Check if required sheets exist
  expectedSheets.forEach(expectedSheet => {
    const found = sheets.find(sheet => 
      sheet.name.toLowerCase().includes(expectedSheet.toLowerCase())
    );
    
    if (!found) {
      const error: ValidationError = {
        id: `missing-sheet-${expectedSheet}`,
        sheet: 'File Structure',
        row: 0,
        column: '',
        field: 'Required Sheet',
        message: `Missing required sheet: ${expectedSheet}`,
        severity: 'error',
        rule: 'SHEET_STRUCTURE_001',
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

  // Sheet-specific validations
  if (sheet.name.toLowerCase().includes('designation mapping')) {
    return validateDesignationMappingSheet(sheet);
  } else if (sheet.name.toLowerCase().includes('survey master')) {
    return validateSurveyMasterSheet(sheet);
  } else if (sheet.name.toLowerCase().includes('question master')) {
    return validateQuestionMasterSheet(sheet);
  } else if (sheet.name.toLowerCase().includes('access sheet')) {
    return validateAccessSheet(sheet);
  }

  return errors;
};

const validateDesignationMappingSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['State', 'Medium', 'Medium_in_english', 'List_of_designations', 'Hierarchical_Level'];
  
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
        rule: 'DESIGNATION_MAPPING_001',
      });
    }
  });

  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2; // +2 because index starts at 0 and we skip header
    
    // Check for empty required fields
    requiredColumns.forEach((col, colIndex) => {
      if (!row[colIndex] || row[colIndex].toString().trim() === '') {
        errors.push({
          id: `empty-field-${sheet.name}-${rowNumber}-${col}`,
          sheet: sheet.name,
          row: rowNumber,
          column: col,
          field: col,
          message: `Required field '${col}' is empty`,
          severity: 'error',
          rule: 'REQUIRED_FIELD_001',
        });
      }
    });

    // Validate Hierarchical_Level is numeric
    const hierarchicalLevelIndex = sheet.headers.indexOf('Hierarchical_Level');
    if (hierarchicalLevelIndex >= 0 && row[hierarchicalLevelIndex]) {
      const value = row[hierarchicalLevelIndex].toString();
      if (isNaN(Number(value)) || !Number.isInteger(Number(value))) {
        errors.push({
          id: `invalid-hierarchy-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Hierarchical_Level',
          field: 'Hierarchical_Level',
          message: 'Hierarchical Level must be a positive integer',
          severity: 'error',
          rule: 'HIERARCHY_LEVEL_001',
        });
      }
    }
  });

  return errors;
};

const validateSurveyMasterSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['Survey ID', 'Survey Name', 'State', 'Medium', 'Medium_in_english', 'Visible_on_report_bot', 'Is Active?'];
  
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
    
    // Check Survey ID uniqueness
    const surveyIdIndex = sheet.headers.indexOf('Survey ID');
    if (surveyIdIndex >= 0 && row[surveyIdIndex]) {
      const surveyId = row[surveyIdIndex].toString().trim();
      if (surveyIds.has(surveyId)) {
        errors.push({
          id: `duplicate-survey-id-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Survey ID',
          field: 'Survey ID',
          message: `Duplicate Survey ID: ${surveyId}`,
          severity: 'error',
          rule: 'SURVEY_ID_UNIQUE_001',
        });
      } else {
        surveyIds.add(surveyId);
      }
      
      // Validate Survey ID format
      if (!/^SVY_[A-Z0-9_]+$/.test(surveyId) || surveyId.length > 25) {
        errors.push({
          id: `invalid-survey-id-format-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Survey ID',
          field: 'Survey ID',
          message: 'Survey ID must follow pattern SVY_[A-Z0-9_]+ and be ≤25 characters',
          severity: 'error',
          rule: 'SURVEY_ID_FORMAT_001',
        });
      }
    }

    // Validate Yes/No fields
    const yesNoFields = ['Visible_on_report_bot', 'Is Active?'];
    yesNoFields.forEach(field => {
      const fieldIndex = sheet.headers.indexOf(field);
      if (fieldIndex >= 0 && row[fieldIndex]) {
        const value = row[fieldIndex].toString().toLowerCase();
        if (!['yes', 'no', 'y', 'n'].includes(value)) {
          errors.push({
            id: `invalid-yes-no-${field}-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: field,
            field: field,
            message: `${field} must be Yes/No or Y/N`,
            severity: 'error',
            rule: 'YES_NO_FIELD_001',
          });
        }
      }
    });
  });

  return errors;
};

const validateQuestionMasterSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['Survey ID', 'Medium', 'Question_ID', 'Question Type', 'Question', 'Question_english', 'Question_Media_Type', 'Is_Mandatory'];
  
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
        rule: 'QUESTION_MASTER_001',
      });
    }
  });

  const questionIds = new Set();
  
  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;
    
    // Check Question_ID format and uniqueness
    const questionIdIndex = sheet.headers.indexOf('Question_ID');
    if (questionIdIndex >= 0 && row[questionIdIndex]) {
      const questionId = row[questionIdIndex].toString().trim();
      
      if (questionIds.has(questionId)) {
        errors.push({
          id: `duplicate-question-id-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_ID',
          field: 'Question_ID',
          message: `Duplicate Question ID: ${questionId}`,
          severity: 'error',
          rule: 'QUESTION_ID_UNIQUE_001',
        });
      } else {
        questionIds.add(questionId);
      }
      
      // Validate Question_ID format (Q followed by number)
      if (!/^Q\d+$/.test(questionId)) {
        errors.push({
          id: `invalid-question-id-format-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_ID',
          field: 'Question_ID',
          message: 'Question ID must follow format Q[1-999] (e.g., Q1, Q2)',
          severity: 'error',
          rule: 'QUESTION_ID_FORMAT_001',
        });
      }
    }

    // Validate Question_Media_Type
    const mediaTypeIndex = sheet.headers.indexOf('Question_Media_Type');
    if (mediaTypeIndex >= 0) {
      const mediaType = row[mediaTypeIndex]?.toString().trim() || '';
      const validMediaTypes = ['Image', 'Video', 'None', ''];
      
      if (mediaType && !validMediaTypes.includes(mediaType)) {
        errors.push({
          id: `invalid-media-type-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Question_Media_Type',
          field: 'Question_Media_Type',
          message: 'Question Media Type must be Image, Video, or None',
          severity: 'error',
          rule: 'MEDIA_TYPE_001',
        });
      }

      // Check if media link is required but missing
      const mediaLinkIndex = sheet.headers.indexOf('Question_Media_Link');
      if (mediaType && mediaType !== 'None' && mediaLinkIndex >= 0) {
        const mediaLink = row[mediaLinkIndex]?.toString().trim() || '';
        if (!mediaLink) {
          errors.push({
            id: `missing-media-link-${rowNumber}`,
            sheet: sheet.name,
            row: rowNumber,
            column: 'Question_Media_Link',
            field: 'Question_Media_Link',
            message: 'Media link is required when Media Type is not None',
            severity: 'error',
            rule: 'MEDIA_LINK_REQUIRED_001',
          });
        }
      }
    }
  });

  return errors;
};

const validateAccessSheet = (sheet: SheetData): ValidationError[] => {
  const errors: ValidationError[] = [];
  const requiredColumns = ['Designation', 'Hierarchical Level', 'State', 'Name', 'Mobile'];
  
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

  // Validate data rows
  sheet.data.slice(1).forEach((row, index) => {
    const rowNumber = index + 2;
    
    // Validate mobile number format
    const mobileIndex = sheet.headers.indexOf('Mobile');
    if (mobileIndex >= 0 && row[mobileIndex]) {
      const mobile = row[mobileIndex].toString().trim();
      if (!/^[6-9]\d{9}$/.test(mobile)) {
        errors.push({
          id: `invalid-mobile-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Mobile',
          field: 'Mobile',
          message: 'Mobile number must be 10 digits starting with 6-9',
          severity: 'error',
          rule: 'MOBILE_FORMAT_001',
        });
      }
    }

    // Validate email format if present
    const emailIndex = sheet.headers.indexOf('Email');
    if (emailIndex >= 0 && row[emailIndex]) {
      const email = row[emailIndex].toString().trim();
      if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        errors.push({
          id: `invalid-email-${rowNumber}`,
          sheet: sheet.name,
          row: rowNumber,
          column: 'Email',
          field: 'Email',
          message: 'Invalid email format',
          severity: 'warning',
          rule: 'EMAIL_FORMAT_001',
        });
      }
    }
  });

  return errors;
};