import React, { useState } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { FileUpload } from './FileUpload';
import { ValidationDashboard } from './ValidationDashboard';
import { parseFile } from '@/utils/fileParser';
import { downloadValidationReport } from '@/utils/exportReport';
import { FileUploadState } from '@/types/validation';
import { Loader2, FileCheck, ArrowLeft } from 'lucide-react';
import { toast } from '@/hooks/use-toast';

export const FMBValidator: React.FC = () => {
  const [state, setState] = useState<FileUploadState>({
    file: null,
    isProcessing: false,
    parsedData: null,
    error: null,
  });

  const handleFileSelect = async (file: File) => {
    setState(prev => ({ ...prev, isProcessing: true, error: null, file }));
    
    try {
      const parsedData = await parseFile(file);
      setState(prev => ({ ...prev, parsedData, isProcessing: false }));
      
      toast({
        title: "File processed successfully",
        description: `Found ${parsedData.validationSummary.totalErrors} errors and ${parsedData.validationSummary.totalWarnings} warnings`,
      });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      setState(prev => ({ ...prev, error: errorMessage, isProcessing: false }));
      
      toast({
        title: "Processing failed",
        description: errorMessage,
        variant: "destructive",
      });
    }
  };


  const handleDownloadReport = () => {
    if (state.parsedData) {
      downloadValidationReport(state.parsedData);
      toast({
        title: "Report downloaded",
        description: "Validation report has been downloaded as CSV",
      });
    }
  };

  const handleReset = () => {
    setState({
      file: null,
      isProcessing: false,
      parsedData: null,
      error: null,
    });
  };

  if (state.isProcessing) {
    return (
      <div className="min-h-screen bg-background">
        <div className="container mx-auto py-8">
          <Card className="max-w-md mx-auto">
            <CardContent className="p-8 text-center">
              <div className="space-y-4">
                <div className="flex justify-center">
                  <div className="p-4 bg-primary/10 rounded-full">
                    <Loader2 className="h-8 w-8 text-primary animate-spin" />
                  </div>
                </div>
                <div>
                  <h3 className="text-lg font-semibold text-foreground">Processing File</h3>
                  <p className="text-muted-foreground mt-1">
                    Analyzing your FMB dump sheet and running validations...
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    );
  }

  if (state.parsedData) {
    return (
      <div className="min-h-screen bg-background">
        <div className="container mx-auto py-8">
          <div className="mb-6">
            <Button 
              variant="outline" 
              onClick={handleReset}
              className="flex items-center space-x-2"
            >
              <ArrowLeft className="h-4 w-4" />
              <span>Upload Another File</span>
            </Button>
          </div>
          <ValidationDashboard 
            parsedFile={state.parsedData} 
            onDownloadReport={handleDownloadReport}
          />
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-background">
      <div className="container mx-auto py-8">
        {/* Header */}
        <div className="text-center mb-8">
          <div className="flex justify-center mb-4">
            <div className="p-4 bg-primary/10 rounded-full">
              <FileCheck className="h-12 w-12 text-primary" />
            </div>
          </div>
          <h1 className="text-4xl font-bold text-foreground mb-4">
            FMB Dump-Sheet Validator
          </h1>
          <p className="text-xl text-muted-foreground max-w-2xl mx-auto">
            Upload your Excel files to validate FMB dump sheets according to comprehensive quality standards
          </p>
        </div>

        {/* Main Upload Area */}
        <div className="max-w-4xl mx-auto">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <FileCheck className="h-5 w-5" />
                <span>File Upload & Validation</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="p-6">
              <FileUpload
                onFileSelect={handleFileSelect}
                isProcessing={state.isProcessing}
                error={state.error}
              />
            </CardContent>
          </Card>

          {/* Information Section */}
          <div className="mt-8 grid md:grid-cols-2 gap-6">
            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Supported Formats</CardTitle>
              </CardHeader>
              <CardContent>
                <ul className="space-y-2 text-sm text-muted-foreground">
                  <li>• Excel files (.xlsx, .xls)</li>
                  <li>• CSV files (.csv)</li>
                </ul>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="text-lg">Validation Checks</CardTitle>
              </CardHeader>
              <CardContent>
                <ul className="space-y-2 text-sm text-muted-foreground">
                  <li>• Required sheet structure validation</li>
                  <li>• Field format and data type checks</li>
                  <li>• Cross-sheet integrity validation</li>
                  <li>• Business logic compliance</li>
                </ul>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
};