import { useState } from 'react';
import { createWorker } from 'tesseract.js';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export default function ImageToExcelConverter() {
  const [image, setImage] = useState(null);
  const [progress, setProgress] = useState(0);
  const [text, setText] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState('');

  const handleImageUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.type.match('image.*')) {
      setError('Please upload an image file');
      return;
    }

    setError('');
    setImage(URL.createObjectURL(file));
    setText('');
    setProgress(0);
  };

const convertImageToText = async () => {
  if (!image) {
    setError('Please upload an image first');
    return;
  }

  setIsProcessing(true);
  setProgress(0);
  setText('');

  try {
    const worker = await createWorker({
      logger: m => {
        if (m.status === 'recognizing text') {
          setProgress(Math.round(m.progress * 100));
        }
      }
    });
    
    // Updated API calls for v4
    await worker.load();
    await worker.loadLanguage('eng');
    await worker.initialize('eng');
    
    const { data } = await worker.recognize(image);
    setText(data.text);
    await worker.terminate();
  } catch (err) {
    setError('Error processing image: ' + err.message);
  } finally {
    setIsProcessing(false);
  }
};
const exportToExcel = () => {
  console.log(text)
  if (!text) {
    setError('No text extracted. Please convert an image first.');
    return;
  }

  try {
    // 1. ENHANCED TEXT PARSING SPECIFIC TO YOUR FORMAT
    const parseTextData = (text) => {
      // Helper function to extract values between labels
      const extractValue = (label, terminator = /[,\n]/) => {
        const pattern = new RegExp(`${label}[\\s:-]*([^${terminator}]+)`);
        const match = text.match(pattern);
        return match ? match[1].trim() : '';
      };

      // Convert written currency to number (e.g., "Three Billion..." → 3628457456.95)
      const parseCurrency = (currencyText) => {
        if (!currencyText) return 0;
        
        // Check if already in numeric format
        const numericMatch = currencyText.match(/(\d[\d,.]*)/);
        if (numericMatch) {
          return parseFloat(numericMatch[0].replace(/,/g, ''));
        }
        
        // Handle written numbers
        const writtenNumbers = {
          zero: 0, one: 1, two: 2, three: 3, four: 4, five: 5,
          six: 6, seven: 7, eight: 8, nine: 9, ten: 10,
          eleven: 11, twelve: 12, thirteen: 13, fourteen: 14, fifteen: 15,
          twenty: 20, thirty: 30, forty: 40, fifty: 50,
          hundred: 100, thousand: 1000, million: 1000000, billion: 1000000000
        };

        let total = 0;
        let current = 0;
        
        currencyText.toLowerCase()
          .replace(/dollars and cents/g, '')
          .split(/[\s-]+/)
          .forEach(word => {
            const num = writtenNumbers[word];
            if (num !== undefined) {
              if (num >= 100) {
                current *= num;
                total += current;
                current = 0;
              } else {
                current += num;
              }
            }
          });
        
        return total + current;
      };

      // Parse percentage values
      const parsePercentage = (percentText) => {
        const match = percentText.match(/(\d+\.?\d*)/);
        return match ? parseFloat(match[0]) : 0;
      };

      // Extract all fields
      return {
        refNumber: extractValue('Customer Reference Number'),
        customerName: extractValue('Customer Name'),
        cityState: extractValue('City State'),
        purchaseValue: parseCurrency(extractValue('Purchase Value \\(USD\\)')),
        downPayment: parsePercentage(extractValue('Down Payment')),
        loanYears: extractValue('Loan Period').match(/\d+/)?.[0] || '',
        annualInterest: parsePercentage(extractValue('Annual Interest')),
        purchaseValueReduction: parsePercentage(extractValue('Purchase Value Reduction')),
        monthlyPrincipalReduction: parsePercentage(extractValue('Monthly Principal Reduction')),
        totalInterestReduction: parsePercentage(extractValue('Total Interest Reduction')),
        guarantorName: extractValue('Guarantor Name'),
        guarantorRef: extractValue('Guarantor Reference Number')
      };
    };

    // 2. PARSE THE EXTRACTED TEXT
    const data = parseTextData(text);
    console.log('Parsed Data:', data);

    // 3. PREPARE EXCEL DATA STRUCTURE
    const headers = [
      "Customer Reference Number",
      "Customer Name",
      "City State",
      "Purchase Value (USD)",
      "Down Payment (%)",
      "Loan Period (Years)",
      "Annual Interest (%)",
      "Purchase Value Reduction (%)",
      "Monthly Principal Reduction (%)",
      "Total Interest Reduction (%)",
      "Guarantor Name",
      "Guarantor Reference Number"
    ];

    const excelData = [
      headers,
      [
        data.refNumber,
        data.customerName,
        data.cityState,
        data.purchaseValue,
        data.downPayment,
        data.loanYears,
        data.annualInterest,
        data.purchaseValueReduction,
        data.monthlyPrincipalReduction,
        data.totalInterestReduction,
        data.guarantorName,
        data.guarantorRef
      ]
    ];

    // 4. CREATE EXCEL WORKSHEET WITH FORMATTING
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    
    // Set column widths
    ws['!cols'] = [
      { wch: 25 }, { wch: 25 }, { wch: 20 },
      { wch: 25 }, { wch: 15 }, { wch: 15 },
      { wch: 15 }, { wch: 20 }, { wch: 20 },
      { wch: 20 }, { wch: 20 }, { wch: 25 }
    ];

    // Format currency and percentages
    const formatCell = (row, col, format) => {
      const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
      if (ws[cellRef]) ws[cellRef].z = format;
    };

    // Format purchase value as currency
    formatCell(1, 3, '"$"#,##0.00');
    
    // Format all percentage columns
    for (let col = 4; col <= 9; col++) {
      formatCell(1, col, '0.00"%"');
    }

    // 5. GENERATE AND DOWNLOAD EXCEL FILE
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Customer Data");
    
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    XLSX.writeFile(wb, `Customer_Data_${timestamp}.xlsx`);

    console.log('Excel file generated successfully!');

  } catch (err) {
    console.error('Excel Export Error:', err);
    setError(`Failed to generate Excel: ${err.message}`);
  }
};


  return (
    <div className="max-w-2xl mx-auto p-6 bg-white rounded-lg shadow-md">
      <h1 className="text-2xl font-bold text-gray-800 mb-6">Image to Excel Converter</h1>
      
      {error && (
        <div className="mb-4 p-3 bg-red-100 text-red-700 rounded-md">
          {error}
        </div>
      )}

      <div className="mb-6">
        <label className="block text-sm font-medium text-gray-700 mb-2">
          Upload Image
        </label>
        <input
          type="file"
          accept="image/*"
          onChange={handleImageUpload}
          className="block w-full text-sm text-gray-500
            file:mr-4 file:py-2 file:px-4
            file:rounded-md file:border-0
            file:text-sm file:font-semibold
            file:bg-blue-50 file:text-blue-700
            hover:file:bg-blue-100"
        />
      </div>

      {image && (
        <div className="mb-6">
          <h2 className="text-lg font-medium text-gray-700 mb-2">Preview</h2>
          <img
            src={image}
            alt="Uploaded preview"
            className="max-h-60 rounded-md border border-gray-200"
          />
        </div>
      )}

      {isProcessing && (
        <div className="mb-6">
          <div className="flex justify-between mb-1">
            <span className="text-sm font-medium text-gray-700">Processing... {progress}%</span>
          </div>
          <div className="w-full bg-gray-200 rounded-full h-2.5">
            <div
              className="bg-blue-600 h-2.5 rounded-full"
              style={{ width: `${progress}%` }}
            ></div>
          </div>
        </div>
      )}

      {text && (
        <div className="mb-6">
          <h2 className="text-lg font-medium text-gray-700 mb-2">Extracted Text</h2>
          <div className="p-4 bg-gray-50 rounded-md border border-gray-200 max-h-40 overflow-y-auto">
            {text.split('\n').map((line, i) => (
              <p key={i}>{line}</p>
            ))}
          </div>
        </div>
      )}

      <div className="flex space-x-4">
        <button
          onClick={convertImageToText}
          disabled={!image || isProcessing}
          className={`px-4 py-2 rounded-md text-white font-medium 
            ${!image || isProcessing ? 'bg-blue-300 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
        >
          {isProcessing ? 'Processing...' : 'Extract Text'}
        </button>

        {/* <button
          onClick={exportToExcel}
          disabled={!text}
          className={`px-4 py-2 rounded-md text-white font-medium 
            ${!text ? 'bg-green-300 cursor-not-allowed' : 'bg-green-600 hover:bg-green-700'}`}
        >
          Export to Excel
        </button> */}
      </div>
    </div>
    
  );
}