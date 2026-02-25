import React, { useState, useEffect, useRef } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import type { IField, ITable } from '@lark-base-open/js-sdk';
import { Button, Toast, Upload, Typography, Card, Space, Modal, TextArea, Select } from '@douyinfe/semi-ui';
import { IconUpload, IconFile, IconHelpCircle, IconDownload } from '@douyinfe/semi-icons';
import PizZip from 'pizzip';
import saveAs from 'file-saver';
import { renderAsync } from 'docx-preview';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

const { Title, Text, Paragraph } = Typography;

// Helper to format cell value to string
const formatCellValue = (val: any): string => {
    if (val === null || val === undefined) return '';

    // Handle timestamps (13 digits)
    // Heuristic: Check if it's a number or string looking like a ms timestamp (e.g. > year 2000)
    const isTimestamp = (v: any) => {
        const num = Number(v);
        // 946684800000 is year 2000
        return !isNaN(num) && num > 946684800000 && String(num).length === 13;
    };

    if (isTimestamp(val)) {
        try {
            const date = new Date(Number(val));
            return date.toLocaleString('zh-CN', { hour12: false });
        } catch (e) {
            // ignore
        }
    }

    if (typeof val === 'string' || typeof val === 'number' || typeof val === 'boolean') return String(val);
    
    // Handle Arrays (MultiSelect, User, etc.)
    if (Array.isArray(val)) {
        return val.map(item => {
            if (typeof item === 'string') return item;
            if (item && typeof item === 'object') {
                // Common properties for name/text
                return item.name || item.text || item.fullAddress || JSON.stringify(item);
            }
            return String(item);
        }).join(', ');
    }
    
    // Handle Objects (SingleSelect, User, Location, etc.)
    if (typeof val === 'object') {
        return val.name || val.text || val.fullAddress || val.value || JSON.stringify(val);
    }
    
    return String(val);
};

export default function App() {
  const [table, setTable] = useState<ITable | null>(null);
  const [fields, setFields] = useState<IField[]>([]);
  const [attachmentFields, setAttachmentFields] = useState<{ label: string, value: string }[]>([]);
  const [selectedAttachFieldId, setSelectedAttachFieldId] = useState<string>('');
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [generatedPdf, setGeneratedPdf] = useState<Blob | null>(null);
  const [generatedName, setGeneratedName] = useState<string>('');
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');
  const [debugData, setDebugData] = useState<string>('');
  const [showDebug, setShowDebug] = useState(false);
  
  // Hidden container for rendering docx to generate PDF
  const previewRef = useRef<HTMLDivElement>(null);

  const init = async () => {
    try {
      const selection = await bitable.base.getSelection();
      if (selection.tableId) {
          const table = await bitable.base.getTableById(selection.tableId);
          setTable(table);
          const fieldList = await table.getFieldList();
          setFields(fieldList);
          
          // Get attachment fields
          const attachFields = await table.getFieldListByType(17);
          const attachOptions = await Promise.all(attachFields.map(async f => ({
              label: await f.getName(),
              value: f.id
          })));
          setAttachmentFields(attachOptions);
          
          if (attachOptions.length > 0) {
              setSelectedAttachFieldId(attachOptions[0].value);
          }
          
          Toast.success('数据已刷新，检测到 ' + fieldList.length + ' 个字段');
      }
    } catch (e) {
      console.error(e);
      setStatus('初始化失败，请在多维表格中运行');
    }
  };

  useEffect(() => {
    init();
  }, []);

  const handleFileUpload = (files: any) => {
    return false; // Prevent auto upload
  };

  const onFileChange = (info: any) => {
      if (info.fileList && info.fileList.length > 0) {
          const file = info.fileList[0].fileInstance || info.fileList[0];
          setTemplateFile(file);
      } else {
          setTemplateFile(null);
      }
  };

  const generateAndExport = async () => {
    if (!templateFile || !table) {
      Toast.error('请先选择模板和数据表');
      return;
    }

    setLoading(true);
    setStatus('正在获取数据...');
    
    try {
      // 1. Get current record data
      const selection = await bitable.base.getSelection();
      if (!selection.recordId) {
        Toast.error('请先选择一行记录');
        setLoading(false);
        return;
      }
      
      const recordData: Record<string, any> = {};
      
      for (const field of fields) {
        const name = await field.getName();
        // Fallback to getCellValue and manual formatting if getCellString fails
        try {
            // Try to get raw value
            const val = await table.getCellValue(field.id, selection.recordId);
            recordData[name] = formatCellValue(val);
        } catch (e) {
            console.warn(`Failed to get value for field ${name}`, e);
            recordData[name] = '';
        }
      }
      
      // Store debug data
      setDebugData(JSON.stringify(recordData, null, 2));

      setStatus('正在生成Word文档...');
      
      // 2. Read template and render
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
            const content = e.target?.result;
            if (!content) return;

            const zip = new PizZip(content as string | ArrayBuffer);
            
            // --- Custom "Python-like" Replacement Logic ---
            // Instead of using docxtemplater (which is strict), we manually parse the XML
            // This mimics the user's Python script behavior: iterating text nodes and replacing placeholders.
            
            const xmlFile = "word/document.xml";
            if (zip.file(xmlFile)) {
                const xmlStr = zip.file(xmlFile).asText();
                const parser = new DOMParser();
                const xmlDoc = parser.parseFromString(xmlStr, "text/xml");
                const texts = xmlDoc.getElementsByTagName("w:t");
                
                let replacedCount = 0;
                
                for (let i = 0; i < texts.length; i++) {
                    const node = texts[i];
                    let text = node.textContent || '';
                    
                    // Regex to match {{key}} or {key}
                    // Captures the key inside the braces
                    const regex = /\{+([^{}]+)\}+/g;
                    
                        if (text.match(regex)) {
                            // Helper to find value case-insensitively
                            const findValue = (k: string) => {
                                // 1. Try exact match
                                let v = recordData[k];
                                if (v !== undefined) return v;

                                // 2. Normalize key for fuzzy match (remove non-alphanumeric/Chinese)
                                // e.g. "A= 一行一列" -> "一行一列", "🔒 一行一列" -> "一行一列"
                                const normalize = (str: string) => str.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '').toLowerCase();
                                const normalizedK = normalize(k);

                                const foundKey = Object.keys(recordData).find(key => normalize(key) === normalizedK);
                                if (foundKey) return recordData[foundKey];
                                
                                return undefined;
                            };

                            const newText = text.replace(regex, (match, key) => {
                                // Trim whitespace from key just in case
                                key = key.trim();
                                
                                let val = findValue(key);
                                
                                if (val !== undefined) {
                                    replacedCount++;
                                    return String(val);
                                } else {
                                    console.warn(`Placeholder not found: ${key}`);
                                    return match; // Keep original if not found
                                }
                            });
                            node.textContent = newText;
                        }
                    }
                
                console.log(`Replaced ${replacedCount} placeholders.`);
                
                // Serialize back to XML
                const serializer = new XMLSerializer();
                const newXml = serializer.serializeToString(xmlDoc);
                zip.file(xmlFile, newXml);
            } else {
                throw new Error("Invalid docx: missing word/document.xml");
            }
            // ---------------------------------------------

            const docxBlob = zip.generate({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });

            // 3. Convert to PDF
            setStatus('正在转换为PDF (这可能需要几秒钟)...');
            
            // Render DOCX to hidden div
            if (previewRef.current) {
                previewRef.current.innerHTML = '';
                await renderAsync(docxBlob, previewRef.current, previewRef.current, {
                    inWrapper: false,
                    ignoreWidth: false,
                    ignoreHeight: false,
                    ignoreFonts: false,
                    breakPages: true,
                    ignoreLastRenderedPageBreak: false,
                    experimental: false,
                    trimXmlDeclaration: true,
                    useBase64URL: false,
                    renderChanges: false,
                    debug: false,
                });
                
                // Wait a bit for images/fonts to render
                await new Promise(r => setTimeout(r, 1000));
                
                // Use html2canvas to capture
                const canvas = await html2canvas(previewRef.current, {
                    scale: 2, // Higher quality
                    useCORS: true
                });
                
                const imgData = canvas.toDataURL('image/png');
                const pdf = new jsPDF({
                    orientation: 'p',
                    unit: 'mm',
                    format: 'a4'
                });
                
                const imgProps = pdf.getImageProperties(imgData);
                const pdfWidth = pdf.internal.pageSize.getWidth();
                const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
                
                pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
                const pdfBlob = pdf.output('blob');
                
                // 4. Upload to Lark
                setStatus('正在上传PDF到多维表格...');
                
                const fileName = `Generated_${selection.recordId}.pdf`;
                setGeneratedName(fileName);
                setGeneratedPdf(pdfBlob);

                if (!selectedAttachFieldId) {
                    Toast.warning('请先选择一个附件字段');
                    saveAs(pdfBlob, fileName);
                    return;
                }

                // Get field name for display
                const selectedOption = attachmentFields.find(f => f.value === selectedAttachFieldId);
                const attachFieldName = selectedOption ? selectedOption.label : '未知字段';

                const file = new File([pdfBlob], fileName, { type: 'application/pdf' });
                
                // Upload file
                setStatus('正在上传文件内容...');
                const tokens = await bitable.base.batchUploadFile([file]);
                if (!tokens || tokens.length === 0) {
                    throw new Error('文件上传失败，未能获取token');
                }
                
                const newAttachment = {
                    token: tokens[0],
                    name: fileName,
                    type: 'application/pdf',
                    timeStamp: Date.now()
                };
                
                // Get current attachments to append
                let currentVal: any[] = [];
                try {
                    const rawVal = await table.getCellValue(selectedAttachFieldId, selection.recordId);
                    if (Array.isArray(rawVal)) {
                        // Filter out any invalid items just in case
                        currentVal = rawVal.filter(item => item && item.token);
                    }
                } catch (e) {
                    console.warn("Failed to get current attachments", e);
                }
                
                setStatus(`正在回写到字段 "${attachFieldName}"...`);
                const finalAttachments = [...currentVal, newAttachment];
                console.log('Writing attachments:', finalAttachments);

                await table.setCellValue(selectedAttachFieldId, selection.recordId, finalAttachments);
                
                // --- Verification Step ---
                setStatus('正在验证回写结果...');
                await new Promise(r => setTimeout(r, 1000)); // wait a bit
                const verifyVal = await table.getCellValue(selectedAttachFieldId, selection.recordId);
                let verified = false;
                if (Array.isArray(verifyVal)) {
                    verified = verifyVal.some((item: any) => item.token === tokens[0]);
                }
                
                if (verified) {
                    Toast.success({ content: `成功！PDF已上传到字段【${attachFieldName}】`, duration: 5 });
                } else {
                    console.error('Verify failed. Expected token:', tokens[0], 'Got:', verifyVal);
                    Toast.warning({ content: `警告：似乎未能写入成功，请尝试手动下载。`, duration: 8 });
                    // Auto download as fallback
                    saveAs(pdfBlob, fileName);
                }
            }
        } catch (err: any) {
            console.error(err);
            Toast.error({ content: '处理失败: ' + err.message, duration: 5 });
        } finally {
            setLoading(false);
            setStatus('');
        }
      };
      reader.readAsArrayBuffer(templateFile);
      
    } catch (err: any) {
      console.error(err);
      Toast.error('发生错误: ' + err.message);
      setLoading(false);
    }
  };

  return (
    <div style={{ padding: 20, maxWidth: 600, margin: '0 auto' }}>
      <Title heading={3} style={{ marginBottom: 20 }}>多维表格排版打印 <Text type="secondary" size="small">(v2.4)</Text></Title>
      
      <Space direction="vertical" style={{ width: '100%' }} spacing="medium">
        <Card>
          <Title heading={5}>1. 准备工作</Title>
          <Space>
            <Button onClick={init} size="small" type="tertiary">刷新表格数据</Button>
            <Text>
                请选择目标附件字段：
            </Text>
            <Select 
                optionList={attachmentFields} 
                value={selectedAttachFieldId} 
                onChange={(v) => setSelectedAttachFieldId(v as string)}
                style={{ width: 150 }}
                placeholder="选择附件字段"
            />
          </Space>
        </Card>

        <Card>
            <Title heading={5} style={{ marginBottom: 10 }}>2. 上传Word模板 (.docx)</Title>
            <Upload
                action=""
                beforeUpload={handleFileUpload}
                onChange={onFileChange}
                limit={1}
                fileList={templateFile ? [{ uid: '1', name: templateFile.name, status: 'success', size: templateFile.size, type: templateFile.type }] : []}
                onRemove={() => setTemplateFile(null)}
                dragMainText="点击或拖拽上传文件"
                dragSubText="支持 .docx 格式"
            >
                {!templateFile && (
                    <div style={{ padding: 20, border: '1px dashed #ccc', borderRadius: 4, textAlign: 'center', cursor: 'pointer' }}>
                        <IconUpload size="large" />
                        <div style={{ marginTop: 8 }}>点击选择模板文件</div>
                    </div>
                )}
            </Upload>
            <div style={{ marginTop: 10 }}>
                <Text type="secondary">
                    模板说明：使用 <Text code>{`{{字段名}}`}</Text> 作为占位符。
                </Text>
            </div>
        </Card>

        <Card>
            <Title heading={5} style={{ marginBottom: 10 }}>3. 生成与导出</Title>
            <Button 
                theme="solid" 
                type="primary" 
                onClick={generateAndExport} 
                loading={loading} 
                disabled={!templateFile}
                block
                size="large"
            >
                {loading ? status || '处理中...' : '生成PDF并回写到附件'}
            </Button>
            {status && <Text style={{ display: 'block', marginTop: 10, textAlign: 'center' }}>{status}</Text>}
            
            <div style={{ marginTop: 10, textAlign: 'right' }}>
                <Space>
                    {generatedPdf && (
                        <Button 
                            type="secondary"
                            icon={<IconDownload />}
                            size="small"
                            onClick={() => saveAs(generatedPdf, generatedName)}
                        >
                            下载PDF
                        </Button>
                    )}
                    <Button 
                        type="tertiary" 
                        icon={<IconHelpCircle />} 
                        size="small"
                        onClick={() => setShowDebug(true)}
                    >
                        查看调试数据
                    </Button>
                </Space>
            </div>
        </Card>
      </Space>

      <Modal
        title="调试数据 (Record Data)"
        visible={showDebug}
        onOk={() => setShowDebug(false)}
        onCancel={() => setShowDebug(false)}
        width={600}
      >
        <Paragraph>
            以下是读取到的当前行数据，请确保Word模板中的 <Text code>{`{{keys}}`}</Text> 与下方的 Key 一致。
        </Paragraph>
        <TextArea 
            value={debugData || '暂无数据，请先点击生成按钮'} 
            autosize 
            readonly 
            style={{ fontFamily: 'monospace', fontSize: 12, minHeight: 200 }} 
        />
      </Modal>

      {/* Hidden container for rendering */}
      <div style={{ position: 'absolute', left: '-9999px', top: 0, width: '210mm', background: 'white' }}>
        <div ref={previewRef}></div>
      </div>
    </div>
  );
}
