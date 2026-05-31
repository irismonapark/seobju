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
import TransformIcon from '@mui/icons-material/Transform';
import DownloadIcon from '@mui/icons-material/Download';
import FileUpload from './FileUpload';
import { convertWorkRecordToInvoice, previewWorkRecordEmployees } from '../utils/converter1';
import { downloadBlob, FileValidationError, getInvoiceDownloadFileName } from '../utils/fileProcessor';

export default function System1() {
  const [file, setFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [resultBlob, setResultBlob] = useState<Blob | null>(null);
  const [previewNames, setPreviewNames] = useState<string[]>([]);

  const handleFileSelect = async (selected: File | null) => {
    setFile(selected);
    setError(null);
    setSuccess(false);
    setResultBlob(null);
    setPreviewNames([]);

    if (!selected) return;

    try {
      const names = await previewWorkRecordEmployees(selected);
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

  const handleConvert = async () => {
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);

    try {
      const blob = await convertWorkRecordToInvoice(file);
      setResultBlob(blob);
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

  const handleDownload = () => {
    if (!resultBlob || !file) return;
    downloadBlob(resultBlob, getInvoiceDownloadFileName(file.name));
  };

  return (
    <Card>
      <CardContent sx={{ p: { xs: 2, sm: 3 } }}>
        <Typography variant="h6" gutterBottom color="primary">
          [시스템 1] 근무현황 → 청구서
        </Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>
          포코스 근무현황 파일(1번)을 업로드하여 청구서(2번) 형식으로 변환합니다.
        </Typography>

        <FileUpload
          label="파일 선택 (1번 파일: 근무현황)"
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

        <Button
          variant="contained"
          startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <TransformIcon />}
          onClick={handleConvert}
          disabled={!file || loading || !!error}
          fullWidth
          sx={{ mb: 2 }}
        >
          {loading ? '변환 중...' : '청구서로 변환'}
        </Button>

        {success && (
          <Alert
            severity="success"
            action={
              <Button color="inherit" size="small" startIcon={<DownloadIcon />} onClick={handleDownload}>
                다운로드
              </Button>
            }
          >
            변환이 완료되었습니다. 다운로드를 시작합니다.
          </Alert>
        )}
      </CardContent>
    </Card>
  );
}
