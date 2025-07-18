import React, { useCallback, useState } from 'react';
import { useDropzone } from 'react-dropzone';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Card, CardContent } from '@/components/ui/card';
import { Upload, FileSpreadsheet, Link2, AlertCircle } from 'lucide-react';
import { Alert, AlertDescription } from '@/components/ui/alert';

interface FileUploadProps {
  onFileSelect: (file: File) => void;
  onGoogleSheetUrl: (url: string) => void;
  isProcessing: boolean;
  error: string | null;
}

export const FileUpload: React.FC<FileUploadProps> = ({
  onFileSelect,
  onGoogleSheetUrl,
  isProcessing,
  error,
}) => {
  const [googleSheetUrl, setGoogleSheetUrl] = useState('');

  const onDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      onFileSelect(acceptedFiles[0]);
    }
  }, [onFileSelect]);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
      'text/csv': ['.csv'],
    },
    multiple: false,
    disabled: isProcessing,
  });

  const handleGoogleSheetSubmit = () => {
    if (googleSheetUrl.trim()) {
      onGoogleSheetUrl(googleSheetUrl.trim());
    }
  };

  return (
    <div className="space-y-6">
      {error && (
        <Alert variant="destructive">
          <AlertCircle className="h-4 w-4" />
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      )}

      {/* File Upload Area */}
      <Card>
        <CardContent className="p-8">
          <div
            {...getRootProps()}
            className={`
              border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors
              ${isDragActive ? 'border-primary bg-primary/5' : 'border-muted-foreground/25'}
              ${isProcessing ? 'opacity-50 cursor-not-allowed' : 'hover:border-primary hover:bg-primary/5'}
            `}
          >
            <input {...getInputProps()} />
            <div className="space-y-4">
              <div className="flex justify-center">
                <div className="p-4 bg-primary/10 rounded-full">
                  <Upload className="h-8 w-8 text-primary" />
                </div>
              </div>
              <div>
                <h3 className="text-lg font-semibold text-foreground">Upload FMB Dump Sheet</h3>
                <p className="text-muted-foreground mt-1">
                  {isDragActive
                    ? 'Drop the file here...'
                    : 'Drag & drop your Excel or CSV file here, or click to browse'}
                </p>
              </div>
              <div className="flex items-center justify-center space-x-2 text-sm text-muted-foreground">
                <FileSpreadsheet className="h-4 w-4" />
                <span>Supports .xlsx, .xls, .csv files</span>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* OR Divider */}
      <div className="relative">
        <div className="absolute inset-0 flex items-center">
          <span className="w-full border-t border-border" />
        </div>
        <div className="relative flex justify-center text-xs uppercase">
          <span className="bg-background px-2 text-muted-foreground">Or</span>
        </div>
      </div>

      {/* Google Sheets Input */}
      <Card>
        <CardContent className="p-6">
          <div className="space-y-4">
            <div className="flex items-center space-x-2">
              <Link2 className="h-5 w-5 text-primary" />
              <h3 className="text-lg font-semibold">Import from Google Sheets</h3>
            </div>
            <p className="text-sm text-muted-foreground">
              Paste a public Google Sheets URL to import your FMB dump sheet directly
            </p>
            <div className="flex space-x-2">
              <Input
                placeholder="https://docs.google.com/spreadsheets/d/..."
                value={googleSheetUrl}
                onChange={(e) => setGoogleSheetUrl(e.target.value)}
                disabled={isProcessing}
                className="flex-1"
              />
              <Button
                onClick={handleGoogleSheetSubmit}
                disabled={!googleSheetUrl.trim() || isProcessing}
                variant="outline"
              >
                Import
              </Button>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};