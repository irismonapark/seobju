import { useRef } from 'react';
import {
  Box,
  Button,
  Typography,
  Chip,
} from '@mui/material';
import CloudUploadIcon from '@mui/icons-material/CloudUpload';
import InsertDriveFileIcon from '@mui/icons-material/InsertDriveFile';

interface FileUploadProps {
  label: string;
  accept?: string;
  selectedFile: File | null;
  onFileSelect: (file: File | null) => void;
  disabled?: boolean;
}

export default function FileUpload({
  label,
  accept = '.xlsx',
  selectedFile,
  onFileSelect,
  disabled = false,
}: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0] ?? null;
    onFileSelect(file);
  };

  const handleClick = () => {
    inputRef.current?.click();
  };

  return (
    <Box>
      <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
        {label}
      </Typography>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        hidden
        onChange={handleChange}
        disabled={disabled}
      />
      <Button
        variant="outlined"
        startIcon={<CloudUploadIcon />}
        onClick={handleClick}
        disabled={disabled}
        fullWidth
        sx={{ mb: 1.5, py: 1.2 }}
      >
        파일 업로드
      </Button>
      {selectedFile && (
        <Chip
          icon={<InsertDriveFileIcon />}
          label={selectedFile.name}
          color="primary"
          variant="outlined"
          onDelete={() => {
            onFileSelect(null);
            if (inputRef.current) inputRef.current.value = '';
          }}
          sx={{ maxWidth: '100%' }}
        />
      )}
    </Box>
  );
}
