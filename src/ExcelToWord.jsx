import React, { useState } from "react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { PDFDocument } from "pdf-lib";
import { Box, Button, Typography, Dialog, DialogTitle, DialogContent, DialogActions, IconButton, Stepper, Step, StepLabel, LinearProgress } from '@mui/material';
import CloseIcon from '@mui/icons-material/Close';
import ContentCopyIcon from '@mui/icons-material/ContentCopy';

const ExcelToWordAndPDF = () => {
  const [excelFiles, setExcelFiles] = useState([]);
  const [templateFiles, setTemplateFiles] = useState([]);
  const [pdfFiles, setPdfFiles] = useState([]);
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [activeStep, setActiveStep] = useState(0);
  const [completedSteps, setCompletedSteps] = useState([]);
  const [progress, setProgress] = useState(0);
  const [showBatchConvertCode, setShowBatchConvertCode] = useState(false);
  const [openDialog, setOpenDialog] = useState(false);

  const steps = ["อัปโหลดไฟล์ Excel", "อัปโหลดไฟล์เทมเพลต", "แปลงไฟล์ Word"];

  const handleExcelFilesChange = (e) => {
    const files = Array.from(e.target.files);
    setExcelFiles(files);

    const allData = [];
    files.forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const binaryStr = event.target.result;
        const workbook = XLSX.read(binaryStr, { type: "binary" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        allData.push(...worksheet);
        setData(allData);
        setActiveStep(1);
        setCompletedSteps((prev) => [...prev, 0]);
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
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName]
          );

          for (
            let templateIndex = 0;
            templateIndex < templateFiles.length;
            templateIndex++
          ) {
            const templateFile = templateFiles[templateIndex];
            const templateReader = new FileReader();
            templateReader.onload = async (event) => {
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
              setShowBatchConvertCode(true);
              setOpenDialog(true); // เปิด Dialog เมื่อการแปลงเสร็จสิ้น
            };
            templateReader.readAsArrayBuffer(templateFile);
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
    const code = 
    `
Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "C:\Users\User\Desktop\New folder (5)\"
    targetFolder = "C:\Users\User\Desktop\pdf\"

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
    <Box className="container" sx={{ padding: 3 }}>
      <Typography variant="h4" gutterBottom>
        ระบบจัดการเอกสาร
      </Typography>

      <Stepper activeStep={activeStep} sx={{ marginBottom: 3 }}>
        {steps.map((label, index) => (
          <Step key={label} completed={completedSteps.includes(index)}>
            <StepLabel>{label}</StepLabel>
          </Step>
        ))}
      </Stepper>

      <Box sx={{ marginBottom: 2 }}>
        {activeStep === 0 && (
          <Box className="step">
            <Typography variant="h6">ขั้นตอนที่ 1: อัปโหลดไฟล์ Excel</Typography>
            <input
              type="file"
              accept=".xlsx, .xls"
              id="excelFiles"
              onChange={handleExcelFilesChange}
              multiple
              style={{ display: 'none' }}
            />
            <label htmlFor="excelFiles">
              <Button variant="contained" component="span" sx={{ marginTop: 2 }}>
                เลือกไฟล์ Excel
              </Button>
            </label>
            <Typography variant="body2" sx={{ marginTop: 2 }}>
              โปรแกรมจะใช้ข้อมูลจากคอลัมน์ "Name" ในไฟล์ Excel
            </Typography>
          </Box>
        )}

        {activeStep === 1 && (
          <Box className="step">
            <Typography variant="h6">ขั้นตอนที่ 2: อัปโหลดไฟล์เทมเพลต</Typography>
            <input
              type="file"
              accept=".docx"
              id="templateFiles"
              onChange={handleTemplateFilesChange}
              multiple
              style={{ display: 'none' }}
            />
            <label htmlFor="templateFiles">
              <Button variant="contained" component="span" sx={{ marginTop: 2 }}>
                เลือกไฟล์เทมเพลต
              </Button>
            </label>
            <Typography variant="body2" sx={{ marginTop: 2 }}>
              ชื่อจาก Excel จะแทนที่ &#123;&#123;name&#125;&#125; ในไฟล์เทมเพลต
            </Typography>
          </Box>
        )}

        {activeStep === 2 && (
          <Box className="step">
            <Typography variant="h6">ขั้นตอนที่ 3: แปลงไฟล์</Typography>
            {excelFiles.length > 0 && templateFiles.length > 0 && (
              <Button
                variant="contained"
                color="primary"
                onClick={handleConvertToWord}
                disabled={loading}
                sx={{ marginTop: 2 }}
              >
                แปลงเป็นไฟล์ Word
              </Button>
            )}
          </Box>
        )}
      </Box>

      <Box sx={{ display: 'flex', justifyContent: 'space-between', marginBottom: 2 }}>
        <Button
          variant="outlined"
          onClick={handleBack}
          disabled={activeStep === 0}
        >
          ย้อนกลับ
        </Button>
      </Box>

      {loading && (
        <Box sx={{ width: '100%', marginBottom: 2 }}>
          <LinearProgress variant="determinate" value={progress} />
          <Typography variant="body2" sx={{ marginTop: 1 }}>{progress}% เสร็จสิ้น</Typography>
        </Box>
      )}

      {data.length > 0 && (
        <Box className="data-preview" sx={{ marginBottom: 3 }}>
          <Typography variant="h6">ตัวอย่างข้อมูล</Typography>
          <table>
            <thead>
              <tr>
                <th>ลำดับ</th>
                <th>ชื่อ</th>
              </tr>
            </thead>
            <tbody>
              {data.slice(0, 5).map((row, index) => (
                <tr key={index}>
                  <td>{index + 1}</td>
                  <td>{row.Name}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {data.length > 5 && (
            <Typography variant="body2">
              แสดง 5 รายการแรกจากทั้งหมด {data.length} รายการ
            </Typography>
          )}
        </Box>
      )}

      {showBatchConvertCode && (
        <Dialog
          open={openDialog}
          onClose={handleCloseDialog}
          aria-labelledby="customized-dialog-title"
          maxWidth="lg" // เพิ่มความกว้างของ Dialog
          fullWidth // เพิ่มความกว้างให้เต็มพื้นที่
        >
          <DialogTitle id="customized-dialog-title" sx={{ m: 0, p: 2 }}>
            โค้ด VBA สำหรับแปลงไฟล์ Word เป็น PDF
          
            <p>เปิด Microsoft Word</p>
  <p>กด Alt + F11 เพื่อเปิด Visual Basic for Applications (VBA) editor</p>
  <p>ไปที่ Insert - Module เพื่อสร้างโมดูลใหม่</p>
            <IconButton
              aria-label="close"
              onClick={handleCloseDialog}
              sx={{
                position: 'absolute',
                right: 8,
                top: 8,
                color: (theme) => theme.palette.grey[500],
              }}
            >
              <CloseIcon />
            </IconButton>
          </DialogTitle>
          <DialogContent dividers>
            <Typography component="pre" gutterBottom sx={{ overflowX: 'auto' }}>
              {`
Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "C:\Users\User\Desktop\New folder (5)\"
    targetFolder = "C:\Users\User\Desktop\pdf\"

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

`}
            </Typography>
          </DialogContent>
          <DialogActions>
            <Button 
              onClick={handleCopyCode} 
              startIcon={<ContentCopyIcon />}
            >
              คัดลอกโค้ด
            </Button>
            <Button autoFocus onClick={handleCloseDialog}>
              ปิด
            </Button>
          </DialogActions>
        </Dialog>
      )}
    </Box>
  );
};

export default ExcelToWordAndPDF;
