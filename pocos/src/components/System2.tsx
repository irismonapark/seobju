import { useState } from 'react';
import {
  Box,
  Button,
  Card,
  CardContent,
  Typography,
  Alert,
  CircularProgress,
  Chip,
  Stack,
} from '@mui/material';
import PictureAsPdfIcon from '@mui/icons-material/PictureAsPdf';
import TableChartIcon from '@mui/icons-material/TableChart';
import DownloadIcon from '@mui/icons-material/Download';
import FileUpload from './FileUpload';
import { convertInvoiceToPayslip, previewInvoiceEmployees } from '../utils/converter2';
import { downloadZip, FileValidationError } from '../utils/fileProcessor';
import type { GeneratedFile } from '../types';

export default function System2() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [generatedFiles, setGeneratedFiles] = useState<GeneratedFile[]>([]);
  const [previewNames, setPreviewNames] = useState<string[]>([]);
  const [lastFormat, setLastFormat] = useState<'pdf' | 'excel'>('pdf');

  const handleFileSelect = async (selected: File | null) => {
    setFile(selected);
    setError(null);
    setSuccess(false);
    setGeneratedFiles([]);
    setPreviewNames([]);

    if (!selected) return;

    try {
      const names = await previewInvoiceEmployees(selected);
      setPreviewNames(names);
    } catch (err) {
      if (err instanceof FileValidationError) {
        setError(err.message);
      } else if (err instanceof Error) {
        setError(err.message);
      }
      setFile(null);
    }
  };

  const handleConvert = async (format: 'pdf' | 'excel') => {
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);
    setLastFormat(format);

    try {
      const files = await convertInvoiceToPayslip(file, format);
      setGeneratedFiles(files);
      setSuccess(true);
    } catch (err) {
      if (err instanceof FileValidationError) {
        setError(err.message);
      } else if (err instanceof Error) {
        setError(err.message);
      } else {
        setError('변환 중 오류가 발생했습니다.');
      }
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadAll = async () => {
    if (generatedFiles.length === 0) return;
    const ext = lastFormat === 'pdf' ? 'pdf' : 'xlsx';
    const baseName = file?.name.replace(/\.xlsx$/i, '') ?? '청구서';
    await downloadZip(generatedFiles, `${baseName}_급여명세서_${ext}.zip`);
  };

  return (
    <Card>
      <CardContent sx={{ p: { xs: 2, sm: 3 } }}>
        <Typography variant="h6" gutterBottom color="primary">
          [시스템 2] 청구서 → 급여명세서
        </Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>
          포코스 청구서 파일(2번)을 업로드하여 직원별 급여명세서(3번)를 생성합니다.
        </Typography>

        <FileUpload
          label="파일 선택 (2번 파일: 청구서)"
          selectedFile={file}
          onFileSelect={handleFileSelect}
          disabled={loading}
        />

        {previewNames.length > 0 && (
          <Box sx={{ mb: 2 }}>
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              변환 대상 직원 ({previewNames.length}명)
            </Typography>
            <Stack direction="row" sx={{ flexWrap: 'wrap', gap: 0.5 }}>
              {previewNames.slice(0, 10).map((name) => (
                <Chip key={name} label={name} size="small" />
              ))}
              {previewNames.length > 10 && (
                <Chip label={`+${previewNames.length - 10}명`} size="small" />
              )}
            </Stack>
          </Box>
        )}

        {error && (
          <Alert severity="error" sx={{ mb: 2 }}>
            {error}
          </Alert>
        )}

        <Stack direction={{ xs: 'column', sm: 'row' }} spacing={1.5} sx={{ mb: 2 }}>
          <Button
            variant="contained"
            startIcon={
              loading && lastFormat === 'pdf' ? (
                <CircularProgress size={18} color="inherit" />
              ) : (
                <PictureAsPdfIcon />
              )
            }
            onClick={() => handleConvert('pdf')}
            disabled={!file || loading || !!error}
            fullWidth
          >
            {loading && lastFormat === 'pdf' ? '변환 중...' : '명세서 PDF변환'}
          </Button>
          <Button
            variant="outlined"
            startIcon={
              loading && lastFormat === 'excel' ? (
                <CircularProgress size={18} color="inherit" />
              ) : (
                <TableChartIcon />
              )
            }
            onClick={() => handleConvert('excel')}
            disabled={!file || loading || !!error}
            fullWidth
          >
            {loading && lastFormat === 'excel' ? '변환 중...' : '명세서 Excel변환'}
          </Button>
        </Stack>

        {success && generatedFiles.length > 0 && (
          <Alert
            severity="success"
            action={
              <Button
                color="inherit"
                size="small"
                startIcon={<DownloadIcon />}
                onClick={handleDownloadAll}
              >
                ZIP 다운로드
              </Button>
            }
          >
            {generatedFiles.length}명의 명세서 생성 완료!
          </Alert>
        )}
      </CardContent>
    </Card>
  );
}
