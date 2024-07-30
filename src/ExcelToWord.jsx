import React, { useState } from "react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import {
  Box,
  Button,
  Typography,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  IconButton,
  Stepper,
  Step,
  StepLabel,
  LinearProgress,
  Paper,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Container,
  Card,
  CardContent,
  Tooltip,
} from "@mui/material";
import CloseIcon from "@mui/icons-material/Close";
import ContentCopyIcon from "@mui/icons-material/ContentCopy";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";
import ArrowBackIcon from "@mui/icons-material/ArrowBack";
import InfoIcon from "@mui/icons-material/Info";

const ExcelToWordAndPDF = () => {
  // States
  const [excelFiles, setExcelFiles] = useState([]);
  const [templateFiles, setTemplateFiles] = useState([]);
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [activeStep, setActiveStep] = useState(0);
  const [completedSteps, setCompletedSteps] = useState([]);
  const [progress, setProgress] = useState(0);
  const [openDialog, setOpenDialog] = useState(false);
  const [showDialogContent, setShowDialogContent] = useState(false); // New state for additional content
  const [conversionCompleted, setConversionCompleted] = useState(false); 

  const steps = ["อัปโหลดไฟล์ Excel", "อัปโหลดไฟล์ Templates", "แปลงไฟล์ Word"];

  const handleExcelFilesChange = (e) => {
    const files = Array.from(e.target.files);
    setExcelFiles(files);

    const allData = [];
    files.forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const binaryStr = event.target.result;
          const workbook = XLSX.read(binaryStr, { type: "binary" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
          allData.push(...worksheet);
          setData(allData);
          setActiveStep(1);
          setCompletedSteps((prev) => [...prev, 0]);
        } catch (error) {
          console.error("Error reading Excel file", error);
        }
      };
      reader.readAsBinaryString(file);
    });
  };

  const handleTemplateFilesChange = (e) => {
    setTemplateFiles(Array.from(e.target.files));
    setActiveStep(2);
    setCompletedSteps((prev) => [...prev, 1]);
  };

  const handleConvertToWord = async () => {
    if (excelFiles.length > 0 && templateFiles.length > 0) {
      setLoading(true);
      setProgress(0);
      const wordZip = new JSZip();
  
      for (let fileIndex = 0; fileIndex < excelFiles.length; fileIndex++) {
        const file = excelFiles[fileIndex];
        const reader = new FileReader();
        reader.onload = async (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const worksheet = XLSX.utils.sheet_to_json(
              workbook.Sheets[sheetName]
            );
  
            for (let templateIndex = 0; templateIndex < templateFiles.length; templateIndex++) {
              const templateFile = templateFiles[templateIndex];
              const templateReader = new FileReader();
              templateReader.onload = async (event) => {
                try {
                  const content = event.target.result;
  
                  for (const [index, record] of worksheet.entries()) {
                    const zipContent = new PizZip(content);
                    const doc = new Docxtemplater(zipContent, {
                      paragraphLoop: true,
                      linebreaks: true,
                    });
  
                    doc.setData({ name: record.Name });
  
                    try {
                      doc.render();
                      const output = doc.getZip().generate({ type: "blob" });
                      const wordFileName = `${index + 1}. ${record.Name}.docx`;
                      wordZip.file(wordFileName, output);
                    } catch (error) {
                      console.error("Error generating document", error);
                    }
                  }
  
                  const wordZipBlob = await wordZip.generateAsync({ type: "blob" });
                  saveAs(
                    wordZipBlob,
                    `Certificates_Word_${fileIndex + 1}_${templateIndex + 1}.zip`
                  );
                  setLoading(false);
                  setConversionCompleted(true); // Set conversion completed to true
                  setOpenDialog(true);
                } catch (error) {
                  console.error("Error processing template file", error);
                }
              };
              templateReader.readAsArrayBuffer(templateFile);
            }
          } catch (error) {
            console.error("Error reading Excel file", error);
          }
        };
        reader.readAsArrayBuffer(file);
      }
      setCompletedSteps((prev) => [...prev, 2]);
      setActiveStep(3);
    } else {
      alert("กรุณาเลือกไฟล์ Excel และไฟล์เทมเพลต");
    }
  };

  const handleBack = () => {
    setActiveStep((prevActiveStep) => Math.max(prevActiveStep - 1, 0));
  };

  const handleCloseDialog = () => {
    setOpenDialog(false);
  };

  const handleCopyCode = () => {
    const code = `
Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "C:\\Users\\User\\Desktop\\New folder (5)\\"
    targetFolder = "C:\\Users\\User\\Desktop\\pdf\\"

    ' รับไฟล์ .docx ทั้งหมดจากโฟลเดอร์ต้นทาง
    file = Dir(sourceFolder & "*.docx")

    ' ทำการวนลูปเพื่อแปลงไฟล์ทั้งหมด
    Do While file <> ""
        ' เปิดไฟล์เอกสาร
        Set doc = Documents.Open(sourceFolder & file)
        
        ' กำหนดชื่อไฟล์ PDF
        docName = Left(file, Len(file) - 5)
        pdfName = targetFolder & docName & ".pdf"
        
        ' บันทึกเป็น PDF
        doc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=wdExportFormatPDF
        
        ' ปิดเอกสาร
        doc.Close False
        
        ' อ่านไฟล์ถัดไป
        file = Dir
    Loop
    
    ' แสดงข้อความเมื่อเสร็จสิ้น
    MsgBox "การแปลงไฟล์ทั้งหมดเสร็จสิ้น!"
End Sub
`;
    navigator.clipboard.writeText(code);
  };

  return (
    <Container maxWidth="md">
      <Paper elevation={3} sx={{ padding: 4, marginTop: 4 }}>
        <Typography variant="h4" gutterBottom align="center">
          Excel to Word Certificate
        </Typography>

        <Stepper
          activeStep={activeStep}
          alternativeLabel
          sx={{ marginBottom: 4 }}
        >
          {steps.map((label, index) => (
            <Step key={label} completed={completedSteps.includes(index)}>
              <StepLabel>{label}</StepLabel>
            </Step>
          ))}
        </Stepper>

        <Card variant="outlined" sx={{ marginBottom: 3 }}>
          <CardContent>
            {activeStep === 0 && (
              <Box>
                <Typography variant="h6" gutterBottom>
                  ขั้นตอนที่ 1: อัปโหลดไฟล์ Excel
                </Typography>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  id="excelFiles"
                  onChange={handleExcelFilesChange}
                  multiple
                  style={{ display: "none" }}
                />
                <label htmlFor="excelFiles">
                  <Button
                    variant="contained"
                    component="span"
                    startIcon={<CloudUploadIcon />}
                    fullWidth
                  >
                    เลือกไฟล์ Excel
                  </Button>
                </label>
                <Tooltip
                  title="โปรแกรมจะใช้ข้อมูลจากคอลัมน์ 'Name' ในไฟล์ Excel"
                  placement="bottom"
                >
                  <Typography
                    variant="body2"
                    sx={{ marginTop: 2, display: "flex", alignItems: "center" }}
                  >
                    <InfoIcon sx={{ marginRight: 1 }} fontSize="small" />
                    ข้อมูลที่ต้องการ: คอลัมน์ "Name" ในไฟล์ Excel
                  </Typography>
                </Tooltip>
              </Box>
            )}

            {activeStep === 1 && (
              <Box>
                <Typography variant="h6" gutterBottom>
                  ขั้นตอนที่ 2: อัปโหลดไฟล์ Templates
                </Typography>
                <input
                  type="file"
                  accept=".docx"
                  id="templateFiles"
                  onChange={handleTemplateFilesChange}
                  multiple
                  style={{ display: "none" }}
                />
                <label htmlFor="templateFiles">
                  <Button
                    variant="contained"
                    component="span"
                    startIcon={<CloudUploadIcon />}
                    fullWidth
                  >
                    เลือกไฟล์ Certificate templates
                  </Button>
                </label>
                <Tooltip
                  title="ชื่อจาก Excel จะแทนที่ {{name}} ในไฟล์เทมเพลต"
                  placement="bottom"
                >
                  <Typography
                    variant="body2"
                    sx={{ marginTop: 2, display: "flex", alignItems: "center" }}
                  >
                    <InfoIcon sx={{ marginRight: 1 }} fontSize="small" />
                    การแทนที่: {"{name}"} ใน templates จะถูกแทนที่ด้วยชื่อจาก
                    Excel
                  </Typography>
                </Tooltip>
              </Box>
            )}

            {activeStep === 2 && (
              <Box>
                <Typography variant="h6" gutterBottom>
                  ขั้นตอนที่ 3: แปลงไฟล์
                </Typography>
                {excelFiles.length > 0 && templateFiles.length > 0 && (
                  <Button
                    variant="contained"
                    color="primary"
                    onClick={handleConvertToWord}
                    disabled={loading}
                    fullWidth
                  >
                    แปลงเป็นไฟล์ Word
                  </Button>
                )}
              </Box>
            )}
          </CardContent>
        </Card>

        <Box
          sx={{
            display: "flex",
            justifyContent: "space-between",
            marginBottom: 2,
          }}
        >
          <Button
            variant="outlined"
            onClick={handleBack}
            disabled={activeStep === 0}
            startIcon={<ArrowBackIcon />}
          >
            ย้อนกลับ
          </Button>
        </Box>

        {loading && (
          <Box sx={{ width: "100%", marginBottom: 2 }}>
            <LinearProgress variant="determinate" value={progress} />
            <Typography variant="body2" sx={{ marginTop: 1 }}>
              {progress}% เสร็จสิ้น
            </Typography>
          </Box>
        )}

        {data.length > 0 && (
          <Paper elevation={2} sx={{ padding: 2, marginTop: 3 }}>
            <Typography variant="h6" gutterBottom>
              ตัวอย่างข้อมูล
            </Typography>
            <TableContainer>
              <Table size="small">
                <TableHead>
                  <TableRow>
                    <TableCell>ลำดับ</TableCell>
                    <TableCell>ชื่อ</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {data.slice(0, 5).map((row, index) => (
                    <TableRow key={index}>
                      <TableCell>{index + 1}</TableCell>
                      <TableCell>{row.Name}</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>
            {data.length > 5 && (
              <Typography variant="body2" sx={{ marginTop: 1 }}>
                แสดง 5 รายการแรกจากทั้งหมด {data.length} รายการ
              </Typography>
            )}
          </Paper>
        )}

        <Button
          variant="contained"
          color="secondary"
          onClick={() => setShowDialogContent((prev) => !prev)}
          sx={{ marginTop: 3 }}
        >
          {showDialogContent ? "ซ่อนขั้นตอนการใช้งาน VBA" : "แสดงขั้นตอนการใช้งาน VBA"}
        </Button>

        {showDialogContent && (
          <DialogContent dividers sx={{ marginTop: 3 }}>
            <Typography variant="body2" paragraph>
              1. เปิด Microsoft Word
            </Typography>
            <Typography variant="body2" paragraph>
              2. กด Alt + F11 เพื่อเปิด Visual Basic for Applications (VBA) editor
            </Typography>
            <Typography variant="body2" paragraph>
              3. ไปที่ Insert - Module เพื่อสร้างโมดูลใหม่
            </Typography>
            <Typography
              component="pre"
              sx={{ overflowX: "auto", bgcolor: "#f5f5f5", padding: 2 }}
            >
              {`Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "C:\\Users\\User\\Desktop\\New folder (5)\\"
    targetFolder = "C:\\Users\\User\\Desktop\\pdf\\"

    ' รับไฟล์ .docx ทั้งหมดจากโฟลเดอร์ต้นทาง
    file = Dir(sourceFolder & "*.docx")

    ' ทำการวนลูปเพื่อแปลงไฟล์ทั้งหมด
    Do While file <> ""
        ' เปิดไฟล์เอกสาร
        Set doc = Documents.Open(sourceFolder & file)
        
        ' กำหนดชื่อไฟล์ PDF
        docName = Left(file, Len(file) - 5)
        pdfName = targetFolder & docName & ".pdf"
        
        ' บันทึกเป็น PDF
        doc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=wdExportFormatPDF
        
        ' ปิดเอกสาร
        doc.Close False
        
        ' อ่านไฟล์ถัดไป
        file = Dir
    Loop
    
    ' แสดงข้อความเมื่อเสร็จสิ้น
    MsgBox "การแปลงไฟล์ทั้งหมดเสร็จสิ้น!"
End Sub`}
            </Typography>
          </DialogContent>
        )}

        <Dialog open={openDialog} onClose={handleCloseDialog} maxWidth="lg" fullWidth>
          <DialogTitle>
            โค้ด VBA สำหรับแปลงไฟล์ Word เป็น PDF
            <IconButton aria-label="close" onClick={handleCloseDialog} sx={{ position: "absolute", right: 8, top: 8 }}>
              <CloseIcon />
            </IconButton>
          </DialogTitle>
          <DialogContent dividers>
            <Typography variant="body2" paragraph>
              1. เปิด Microsoft Word
            </Typography>
            <Typography variant="body2" paragraph>
              2. กด Alt + F11 เพื่อเปิด Visual Basic for Applications (VBA) editor
            </Typography>
            <Typography variant="body2" paragraph>
              3. ไปที่ Insert - Module เพื่อสร้างโมดูลใหม่
            </Typography>
            <Typography
              component="pre"
              sx={{ overflowX: "auto", bgcolor: "#f5f5f5", padding: 2 }}
            >
              {`Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "C:\\Users\\User\\Desktop\\New folder (5)\\"
    targetFolder = "C:\\Users\\User\\Desktop\\pdf\\"

    ' รับไฟล์ .docx ทั้งหมดจากโฟลเดอร์ต้นทาง
    file = Dir(sourceFolder & "*.docx")

    ' ทำการวนลูปเพื่อแปลงไฟล์ทั้งหมด
    Do While file <> ""
        ' เปิดไฟล์เอกสาร
        Set doc = Documents.Open(sourceFolder & file)
        
        ' กำหนดชื่อไฟล์ PDF
        docName = Left(file, Len(file) - 5)
        pdfName = targetFolder & docName & ".pdf"
        
        ' บันทึกเป็น PDF
        doc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=wdExportFormatPDF
        
        ' ปิดเอกสาร
        doc.Close False
        
        ' อ่านไฟล์ถัดไป
        file = Dir
    Loop
    
    ' แสดงข้อความเมื่อเสร็จสิ้น
    MsgBox "การแปลงไฟล์ทั้งหมดเสร็จสิ้น!"
End Sub`}
            </Typography>
          </DialogContent>
          <DialogActions>
            <Button onClick={handleCopyCode} startIcon={<ContentCopyIcon />}>
              คัดลอกโค้ด
            </Button>
            <Button onClick={handleCloseDialog} color="primary">
              ปิด
            </Button>
          </DialogActions>
        </Dialog>
      </Paper>
    </Container>
  );
};

export default ExcelToWordAndPDF;
