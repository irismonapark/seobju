import { useState } from 'react';
import {
  AppBar,
  Box,
  Container,
  Toolbar,
  Typography,
  Divider,
  Accordion,
  AccordionSummary,
  AccordionDetails,
  IconButton,
  Tooltip,
} from '@mui/material';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import HelpOutlineIcon from '@mui/icons-material/HelpOutlineOutlined';
import RefreshIcon from '@mui/icons-material/Refresh';
import System1 from './components/System1';
import System2 from './components/System2';

function App() {
  const [helpExpanded, setHelpExpanded] = useState(false);
  const [refreshKey, setRefreshKey] = useState(0);

  const handleRefresh = () => {
    setRefreshKey((prev) => prev + 1);
    setHelpExpanded(false);
  };

  return (
    <Box sx={{ minHeight: '100vh', bgcolor: 'background.default' }}>
      <AppBar position="static" elevation={1} sx={{ bgcolor: '#00695C' }}>
        <Toolbar>
          <Typography variant="h6" component="h1" sx={{ flexGrow: 1 }}>
            SJ-포코스 급여관리 시스템
          </Typography>
          <Tooltip title="새로고침">
            <IconButton color="inherit" onClick={handleRefresh} aria-label="새로고침">
              <RefreshIcon />
            </IconButton>
          </Tooltip>
        </Toolbar>
      </AppBar>

      <Container maxWidth="md" sx={{ py: { xs: 2, sm: 4 } }}>
        <Box sx={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          <System1 key={`system1-${refreshKey}`} />
          <Divider />
          <System2 key={`system2-${refreshKey}`} />

          <Accordion
            expanded={helpExpanded}
            onChange={(_, expanded) => setHelpExpanded(expanded)}
            sx={{ borderRadius: 2, '&:before': { display: 'none' } }}
          >
            <AccordionSummary expandIcon={<ExpandMoreIcon />}>
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                <HelpOutlineIcon color="primary" />
                <Typography sx={{ fontWeight: 600 }}>사용방법</Typography>
              </Box>
            </AccordionSummary>
            <AccordionDetails>
              <Typography variant="subtitle2" gutterBottom>
                [시스템 1] 근무현황 → 청구서
              </Typography>
              <Typography variant="body2" sx={{ mb: 2 }}>
                1. &quot;파일 선택&quot;에서 근무현황 파일(1번)을 선택합니다.
                <br />
                2. &quot;청구서로 변환&quot; 버튼을 클릭합니다.
                <br />
                3. 변환이 완료되면 &quot;다운로드&quot;를 클릭하여 청구서를 저장합니다.
              </Typography>

              <Typography variant="subtitle2" gutterBottom>
                [시스템 2] 청구서 → 급여명세서
              </Typography>
              <Typography variant="body2" sx={{ mb: 2 }}>
                1. &quot;파일 선택&quot;에서 청구서 파일(2번)을 선택합니다.
                <br />
                2. PDF 또는 Excel 중 원하는 형식을 선택합니다.
                <br />
                3. 각 직원의 명세서가 ZIP 파일로 다운로드됩니다.
              </Typography>

              <Typography variant="subtitle2" gutterBottom>
                주의사항
              </Typography>
              <Typography variant="body2" color="text.secondary">
                • 파일은 반드시 지정된 형식(.xlsx)이어야 합니다.
                <br />
                • 대용량 파일은 처리에 시간이 걸릴 수 있습니다.
                <br />
                • 다운로드 후 파일명을 변경하지 마세요.
              </Typography>
            </AccordionDetails>
          </Accordion>
        </Box>
      </Container>
    </Box>
  );
}

export default App;
