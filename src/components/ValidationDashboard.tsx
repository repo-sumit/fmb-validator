import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  Download, 
  FileText,
  Sheet
} from 'lucide-react';
import { ParsedFile, ValidationError } from '@/types/validation';

interface ValidationDashboardProps {
  parsedFile: ParsedFile;
  onDownloadReport: () => void;
}

export const ValidationDashboard: React.FC<ValidationDashboardProps> = ({
  parsedFile,
  onDownloadReport,
}) => {
  const { validationSummary, validationErrors, sheets } = parsedFile;

  const getSeverityIcon = (severity: ValidationError['severity']) => {
    switch (severity) {
      case 'error':
        return <XCircle className="h-4 w-4 text-destructive" />;
      case 'warning':
        return <AlertTriangle className="h-4 w-4 text-warning" />;
      default:
        return <CheckCircle2 className="h-4 w-4 text-success" />;
    }
  };

  const getSeverityBadge = (severity: ValidationError['severity']) => {
    switch (severity) {
      case 'error':
        return <Badge variant="destructive">Error</Badge>;
      case 'warning':
        return <Badge className="bg-warning text-warning-foreground">Warning</Badge>;
      default:
        return <Badge className="bg-success text-success-foreground">Info</Badge>;
    }
  };

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-2xl font-bold text-foreground">Validation Results</h2>
          <p className="text-muted-foreground">File: {parsedFile.fileName}</p>
        </div>
        <Button onClick={onDownloadReport} variant="outline" className="flex items-center space-x-2">
          <Download className="h-4 w-4" />
          <span>Download Report</span>
        </Button>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
        <Card>
          <CardContent className="p-6">
            <div className="flex items-center space-x-2">
              <Sheet className="h-5 w-5 text-primary" />
              <div>
                <p className="text-sm font-medium text-muted-foreground">Sheets Processed</p>
                <p className="text-2xl font-bold text-foreground">{sheets.length}</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center space-x-2">
              <XCircle className="h-5 w-5 text-destructive" />
              <div>
                <p className="text-sm font-medium text-muted-foreground">Errors</p>
                <p className="text-2xl font-bold text-destructive">{validationSummary.totalErrors}</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center space-x-2">
              <AlertTriangle className="h-5 w-5 text-warning" />
              <div>
                <p className="text-sm font-medium text-muted-foreground">Warnings</p>
                <p className="text-2xl font-bold text-warning">{validationSummary.totalWarnings}</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <Card>
          <CardContent className="p-6">
            <div className="flex items-center space-x-2">
              <CheckCircle2 className="h-5 w-5 text-success" />
              <div>
                <p className="text-sm font-medium text-muted-foreground">Status</p>
                <p className="text-lg font-semibold">
                  {validationSummary.totalErrors === 0 ? (
                    <span className="text-success">Valid</span>
                  ) : (
                    <span className="text-destructive">Issues Found</span>
                  )}
                </p>
              </div>
            </div>
          </CardContent>
        </Card>
      </div>

      {/* Overall Status Alert */}
      {validationSummary.totalErrors === 0 ? (
        <Alert className="border-success bg-success/5">
          <CheckCircle2 className="h-4 w-4 text-success" />
          <AlertDescription className="text-success">
            All validation checks passed! Your FMB dump sheet is ready for ingestion.
          </AlertDescription>
        </Alert>
      ) : (
        <Alert variant="destructive">
          <XCircle className="h-4 w-4" />
          <AlertDescription>
            Found {validationSummary.totalErrors} error(s) and {validationSummary.totalWarnings} warning(s). 
            Please review and fix the issues before proceeding.
          </AlertDescription>
        </Alert>
      )}

      {/* Detailed Results */}
      <Tabs defaultValue="errors" className="w-full">
        <TabsList className="grid w-full grid-cols-3">
          <TabsTrigger value="errors" className="flex items-center space-x-2">
            <XCircle className="h-4 w-4" />
            <span>Errors ({validationSummary.totalErrors})</span>
          </TabsTrigger>
          <TabsTrigger value="warnings" className="flex items-center space-x-2">
            <AlertTriangle className="h-4 w-4" />
            <span>Warnings ({validationSummary.totalWarnings})</span>
          </TabsTrigger>
          <TabsTrigger value="sheets" className="flex items-center space-x-2">
            <FileText className="h-4 w-4" />
            <span>Sheet Summary</span>
          </TabsTrigger>
        </TabsList>

        <TabsContent value="errors" className="space-y-4">
          {validationErrors.filter(error => error.severity === 'error').length === 0 ? (
            <Card>
              <CardContent className="p-8 text-center">
                <CheckCircle2 className="h-12 w-12 text-success mx-auto mb-4" />
                <p className="text-muted-foreground">No errors found! ✨</p>
              </CardContent>
            </Card>
          ) : (
            <div className="space-y-3">
              {validationErrors
                .filter(error => error.severity === 'error')
                .map((error) => (
                  <Card key={error.id}>
                    <CardContent className="p-4">
                      <div className="flex items-start space-x-3">
                        {getSeverityIcon(error.severity)}
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center space-x-2 mb-1">
                            {getSeverityBadge(error.severity)}
                            <Badge variant="outline">{error.sheet}</Badge>
                            <Badge variant="outline">Row {error.row}</Badge>
                          </div>
                          <p className="text-sm font-medium text-foreground">{error.field}</p>
                          <p className="text-sm text-muted-foreground">{error.message}</p>
                          <p className="text-xs text-muted-foreground mt-1">Rule: {error.rule}</p>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
            </div>
          )}
        </TabsContent>

        <TabsContent value="warnings" className="space-y-4">
          {validationErrors.filter(error => error.severity === 'warning').length === 0 ? (
            <Card>
              <CardContent className="p-8 text-center">
                <CheckCircle2 className="h-12 w-12 text-success mx-auto mb-4" />
                <p className="text-muted-foreground">No warnings found!</p>
              </CardContent>
            </Card>
          ) : (
            <div className="space-y-3">
              {validationErrors
                .filter(error => error.severity === 'warning')
                .map((error) => (
                  <Card key={error.id}>
                    <CardContent className="p-4">
                      <div className="flex items-start space-x-3">
                        {getSeverityIcon(error.severity)}
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center space-x-2 mb-1">
                            {getSeverityBadge(error.severity)}
                            <Badge variant="outline">{error.sheet}</Badge>
                            <Badge variant="outline">Row {error.row}</Badge>
                          </div>
                          <p className="text-sm font-medium text-foreground">{error.field}</p>
                          <p className="text-sm text-muted-foreground">{error.message}</p>
                          <p className="text-xs text-muted-foreground mt-1">Rule: {error.rule}</p>
                        </div>
                      </div>
                    </CardContent>
                  </Card>
                ))}
            </div>
          )}
        </TabsContent>

        <TabsContent value="sheets" className="space-y-4">
          <div className="grid gap-4">
            {sheets.map((sheet, index) => (
              <Card key={index}>
                <CardHeader>
                  <CardTitle className="flex items-center justify-between">
                    <span>{sheet.name}</span>
                    <div className="flex space-x-2">
                      <Badge variant="outline">{sheet.data.length} rows</Badge>
                      <Badge variant="outline">{sheet.headers.length} columns</Badge>
                    </div>
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="flex space-x-4 text-sm">
                    <div className="flex items-center space-x-1">
                      <XCircle className="h-4 w-4 text-destructive" />
                      <span>{validationSummary.errorsBySheet[sheet.name] || 0} errors</span>
                    </div>
                    <div className="flex items-center space-x-1">
                      <AlertTriangle className="h-4 w-4 text-warning" />
                      <span>{validationSummary.warningsBySheet[sheet.name] || 0} warnings</span>
                    </div>
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
        </TabsContent>
      </Tabs>
    </div>
  );
};