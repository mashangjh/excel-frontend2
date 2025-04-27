import { useState } from 'react';
import { Upload, Button, Form, InputNumber, Alert, Progress } from 'antd';
import { read, utils } from 'xlsx';
// 尝试安装类型定义文件，若已安装则可直接使用以下导入
// 若已安装 @types/file-saver，则使用原始导入
import { saveAs } from 'file-saver';
// 若未安装，可在项目中添加一个 .d.ts 文件，在其中添加 declare module 'file-saver';
// 若安装失败，可在项目中添加一个 .d.ts 文件，在其中添加 declare module 'file-saver';
import { DiffOutlined, UploadOutlined } from '@ant-design/icons';
// 引入 message 组件
import { message } from 'antd';

type ExcelData = Array<Record<string, string | number>>;

function indexToColumn(index: number): string {
  let column = '';
  while (index >= 0) {
    column = String.fromCharCode(65 + (index % 26)) + column;
    index = Math.floor(index / 26) - 1;
  }
  return column || 'A';
}


import '@ant-design/v5-patch-for-react-19';

export default function App() {
  const [form] = Form.useForm();
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState(0);
  const [diffResult, setDiffResult] = useState<Array<{
    row: number;
    col: string;
    val1: string | number;
    val2: string | number;
  }>>([]);

  function isRecord(value: unknown): value is Record<string, any> {
    return typeof value === 'object' && value !== null;
  }

  const parseExcel = async (file: File): Promise<Record<string, ExcelData>> => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = read(buffer, { type: 'array', codepage: 65001 });
      const allData: Record<string, ExcelData> = {};
      wb.SheetNames.forEach((sheetName) => {
        const data = utils.sheet_to_json(wb.Sheets[sheetName]);
        console.log(`解析 ${sheetName} 后的Excel数据:`, data);
        if (data.length === 0) {
          console.warn(`警告：解析 ${sheetName} 的Excel数据为空`);
        }

        // 处理特殊字符转义
        allData[sheetName] = data.map((value): Record<string, string | number> => {
          if (isRecord(value)) {
            const newRow: Record<string, string | number> = {};
            for (const [key, val] of Object.entries(value)) {
              // 过滤_EMPTY开头的列名
              if (/^_EMPTY\d*$/.test(key)) continue;

              if (typeof val === 'string') {
                newRow[key] = encodeURIComponent(val);
              } else if (typeof val === 'number') {
                newRow[key] = val;
              } else {
                newRow[key] = String(val);
              }
            }
            return newRow;
          }
          return {};
        });
      });
      return allData;
    } catch (error) {
      if (error instanceof Error && error.message.includes('Invalid codepage')) {
        throw new Error('文件编码异常，请确保使用UTF-8编码保存Excel文件');
      }
      throw error;
    }
  };

  const onFinish = async (values: any) => {
    console.log('开始执行 onFinish 函数，表单值:', values);
    try {
      console.log('开始表单验证');
      await form.validateFields();
      console.log('表单验证通过');
      setLoading(true);
      setProgress(0);

      message.info('开始处理文件，请稍候...');
      console.log('开始解析文件');
      if (!values.file1?.file?.originFileObj || !values.file2?.file?.originFileObj) {
        console.error('文件对象未正确传递');
        message.error('文件对象未正确传递，请重新选择文件');
        setLoading(false);
        return;
      }
      const [file1All, file2All] = await Promise.all([
        parseExcel(values.file1.file.originFileObj),
        parseExcel(values.file2.file.originFileObj)
      ]);

      const allSheetNames = Array.from(new Set([...Object.keys(file1All), ...Object.keys(file2All)]));
      let allDifferences: typeof diffResult = [];

      allSheetNames.forEach((sheetName) => {
        const file1 = file1All[sheetName] || [];
        const file2 = file2All[sheetName] || [];
        console.log(`开始对比 ${sheetName}，基准文件:`, file1, '对比文件:', file2, '文件1长度:', file1.length, '文件2长度:', file2.length);

        const differences: typeof diffResult = [];
        const allHeaders = Array.from(new Set([
          ...file1.flatMap(Object.keys).filter(k => !/^_EMPTY\d*$/.test(k)),
          ...file2.flatMap(Object.keys).filter(k => !/^_EMPTY\d*$/.test(k))
        ])).filter(k => k && !/^_EMPTY\d*$/.test(k));

        file1.forEach((row1, index) => {
          const row2 = file2[index] || {};
          allHeaders.forEach(col => {
            if (compareValues(row1[col], row2[col], values.threshold)) {
              differences.push({ 
                row: index + 2,
                col: `${sheetName}-${col}`,
                val1: row1[col] ?? 'N/A',
                val2: row2[col] ?? 'N/A'
              });
            }
          });
        });

        allDifferences = allDifferences.concat(differences);
      });

      setProgress(100);
      setDiffResult(allDifferences);

      // 生成差异报告
      const report = allDifferences.map(d => {
        const [sheetName, col] = d.col.split('-');
        const colIndex = Object.keys(file1All[sheetName]?.[0] || {}).indexOf(col);
        const colLetter = indexToColumn(colIndex);
        return `工作表: ${sheetName}, 行号: ${d.row}, 列: ${colLetter}, 基准值: ${typeof d.val1 === 'string' ? decodeURIComponent(d.val1).trim(): d.val1}, 对比值: ${typeof d.val2 === 'string' ? decodeURIComponent(d.val2).trim() : d.val2}`;
      }).join('\n');

      const blob = new Blob([report], { type: 'text/plain;charset=utf-8' });
      saveAs(blob, '差异报告.txt');

    } catch (error) {
      console.error('对比出错:', error);
      message.error(`处理失败: ${error instanceof Error ? error.message : '未知错误'}`);
    } finally {
      setLoading(false);
      setProgress(0);
    }
  };

  const compareValues = (val1: any, val2: any, threshold: number) => {
    console.log('开始执行 type:',typeof val1,val1);
    if (typeof val1 === 'number' && typeof val2 === 'number') {
      const diff = Math.abs(val1 - val2);
      const relativeDiff = diff / Math.max(Math.abs(val1), Math.abs(val2));
      console.log('数值对比结果:', {val1, val2, diff, threshold, relativeDiff, result: diff > threshold});
      return diff > threshold;
    }
    if (typeof val1 === 'string' && typeof val2 === 'string') {
      try {
        const decodedVal1 = decodeURIComponent(val1).trim();
        const decodedVal2 = decodeURIComponent(val2).trim();
        const result = decodedVal1 !== decodedVal2;
        console.log('字符串对比结果:', {val1: decodedVal1, val2: decodedVal2, result});
        return result;
      } catch (e) {
        const result = val1 !== val2;
        console.log('字符串对比结果（解码失败，使用原始值）:', {val1, val2, result});
        return result;
      }
    }
    const result = String(val1) !== String(val2);
    console.log('默认类型对比结果:', {val1, val2, result, '类型1': typeof val1, '类型2': typeof val2});
    return result;
  };

  return (
    <div style={{ maxWidth: 800, margin: '20px auto', padding: 20 }}>
      <Form form={form} onFinish={onFinish} layout="vertical">
        <Form.Item label="基准文件" name="file1" rules={[{ required: true }]}>
          <Upload.Dragger 
            accept=".xlsx,.xls" 
            maxCount={1}
            beforeUpload={() => false}
            onChange={({ fileList }) => form.setFieldsValue({ file1: { file: fileList[0] } })} // 添加 onChange 事件
          >
            <p><UploadOutlined /> 点击或拖放基准文件</p>
          </Upload.Dragger>
        </Form.Item>

        <Form.Item label="对比文件" name="file2" rules={[{ required: true }]}>
          <Upload.Dragger 
            accept=".xlsx,.xls" 
            maxCount={1}
            beforeUpload={() => false}
            onChange={({ fileList }) => form.setFieldsValue({ file2: { file: fileList[0] } })} // 添加 onChange 事件
          >
            <p><UploadOutlined /> 点击或拖放对比文件</p>
          </Upload.Dragger>
        </Form.Item>



        <Form.Item 
          label="数字阈值" 
          name="threshold" 
          initialValue={0.01}
          tooltip="允许的数字差异百分比（绝对值）"
          rules={[
            { required: true, message: '请输入阈值' },
            {
              validator: (_, value) => 
                value >= 0 && value <= 1 
                  ? Promise.resolve() 
                  : Promise.reject(new Error('阈值必须在0～1之间'))
            }
          ]}
        >
          <InputNumber 
            step={0.01} 
            min={0} 
            formatter={value => `${value! * 100}%`}
            parser={(value) => {
              try {
                const numValue = parseFloat(value!.replace('%', '')) / 100;
                if (isNaN(numValue)) throw new Error('无效数字格式');
                return numValue;
              } catch (error) {
                return 0;
              }
            }
          }
          />
        </Form.Item>

        <Button
          type="primary"
          htmlType="submit"
          icon={<DiffOutlined />}
          loading={loading}
          style={{ marginBottom: 16 }}
        >
          开始对比
        </Button>

        <Progress
          percent={progress}
          status={loading ? 'active' : undefined}
          style={{ marginBottom: 24 }}
        />

        {diffResult.length > 0 && (
          <Alert
            message={`发现 ${diffResult.length} 处差异`}
            description="差异报告文件已自动下载"
            type="info"
            showIcon
          />
        )}
      </Form>
    </div>
  );
}
