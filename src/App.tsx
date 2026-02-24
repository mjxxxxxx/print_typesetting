import React, { useState, useEffect, useRef } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import type { IField, ITable } from '@lark-base-open/js-sdk';
import { Button, Toast, Upload, Typography, Card, Space, Spin, Modal, TextArea } from '@douyinfe/semi-ui';
import { IconUpload, IconFile, IconHelpCircle } from '@douyinfe/semi-icons';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { renderAsync } from 'docx-preview';
import jsPDF from 'jspdf';
import html2canvas from 'html2canvas';

const { Title, Text, Paragraph } = Typography;

// Helper to format cell value to string
const formatCellValue = (val: any): string => {
    if (val === null || val === undefined) return '';
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
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState('');
  const [debugData, setDebugData] = useState<string>('');
  const [showDebug, setShowDebug] = useState(false);
  
  // Hidden container for rendering docx to generate PDF
  const previewRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const init = async () => {
      try {
        const selection = await bitable.base.getSelection();
        if (selection.tableId) {
            const table = await bitable.base.getTableById(selection.tableId);
            setTable(table);
            const fieldList = await table.getFieldList();
            setFields(fieldList);
        }
      } catch (e) {
        console.error(e);
        setStatus('初始化失败，请在多维表格中运行');
      }
    };
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
            const doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter: () => { return ""; } // return empty string for missing values
            });

            try {
                doc.setData(recordData);
                doc.render();
            } catch (error: any) {
                // Improved error handling
                let errorMsg = error.message;
                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (err: any) {
                        return err.properties.explanation;
                    }).join("\n");
                    errorMsg = `模板错误详情:\n${errorMessages}`;
                }
                // Show debug modal automatically on error
                setShowDebug(true);
                throw new Error(errorMsg);
            }

            const docxBlob = doc.getZip().generate({
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
                
                // Find attachment field
                // Note: field type 17 is Attachment
                const attachmentFields = await table.getFieldListByType(17);
                if (attachmentFields.length === 0) {
                    Toast.warning('未找到附件字段，无法回写。正在下载PDF...');
                    pdf.save(`generated_${selection.recordId}.pdf`);
                } else {
                    const attachField = attachmentFields[0];
                    const fileName = `Generated_${selection.recordId}.pdf`;
                    const file = new File([pdfBlob], fileName, { type: 'application/pdf' });
                    
                    // Upload file
                    const tokens = await bitable.base.batchUploadFile([file]);
                    
                    const newAttachment = {
                        token: tokens[0],
                        name: fileName,
                        type: 'application/pdf',
                        timeStamp: Date.now()
                    };
                    
                    // Get current attachments to append
                    // Use getCellValue for consistency
                    let currentVal: any[] = [];
                    try {
                        const rawVal = await table.getCellValue(attachField.id, selection.recordId);
                        if (Array.isArray(rawVal)) {
                            currentVal = rawVal;
                        }
                    } catch (e) {
                        console.warn("Failed to get current attachments", e);
                    }
                    
                    await table.setCellValue(attachField.id, selection.recordId, [...currentVal, newAttachment]);
                    Toast.success('成功！PDF已生成并上传。');
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
      <Title heading={3} style={{ marginBottom: 20 }}>多维表格排版打印 <Text type="secondary" size="small">(v1.2)</Text></Title>
      
      <Space direction="vertical" style={{ width: '100%' }} spacing="medium">
        <Card>
          <Title heading={5}>1. 准备工作</Title>
          <Text>
            请确保当前多维表格中有一行记录被选中，并且表中有一个附件字段用于接收结果。
          </Text>
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
                <Button 
                    type="tertiary" 
                    icon={<IconHelpCircle />} 
                    size="small"
                    onClick={() => setShowDebug(true)}
                >
                    查看调试数据
                </Button>
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
