import { ParsedFile } from '@/types/validation';

export const downloadValidationReport = (parsedFile: ParsedFile) => {
  const { validationErrors, validationSummary, fileName } = parsedFile;
  
  // Create CSV content
  const headers = ['Sheet', 'Row', 'Column', 'Field', 'Message', 'Severity', 'Rule'];
  const csvRows = [
    headers.join(','),
    ...validationErrors.map(error => [
      `"${error.sheet}"`,
      error.row.toString(),
      `"${error.column}"`,
      `"${error.field}"`,
      `"${error.message.replace(/"/g, '""')}"`,
      error.severity,
      `"${error.rule}"`
    ].join(','))
  ];
  
  // Add summary at the end
  csvRows.push('');
  csvRows.push('SUMMARY');
  csvRows.push(`Total Errors,${validationSummary.totalErrors}`);
  csvRows.push(`Total Warnings,${validationSummary.totalWarnings}`);
  csvRows.push('');
  csvRows.push('Errors by Sheet');
  Object.entries(validationSummary.errorsBySheet).forEach(([sheet, count]) => {
    csvRows.push(`"${sheet}",${count}`);
  });
  csvRows.push('');
  csvRows.push('Warnings by Sheet');
  Object.entries(validationSummary.warningsBySheet).forEach(([sheet, count]) => {
    csvRows.push(`"${sheet}",${count}`);
  });
  
  const csvContent = csvRows.join('\n');
  
  // Create and download file
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  link.setAttribute('href', url);
  link.setAttribute('download', `${fileName.replace(/\.[^/.]+$/, '')}_validation_report.csv`);
  link.style.visibility = 'hidden';
  
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
};