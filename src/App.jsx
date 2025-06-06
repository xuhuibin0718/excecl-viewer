import { useState } from 'react';
import ExcelPreview from './components/ExcelPreview/ExcelPreview';
import './App.css';

function App() {
  const [fileStream, setFileStream] = useState(null);
  const [fileName, setFileName] = useState('');
  const [showPreview, setShowPreview] = useState(false);

  const handleFileChange = (event) => {
    const file = event.target.files[0];
    if (file) {
      setFileName(file.name);
      
      // Read file as ArrayBuffer
      const reader = new FileReader();
      reader.onload = (e) => {
        setFileStream(e.target.result);
        setShowPreview(true);
      };
      reader.readAsArrayBuffer(file);
    }
  };

  return (
    <div className="app">
      <div className="upload-container">
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          id="file-upload"
          className="file-input"
        />
        <label htmlFor="file-upload" className="upload-button">
          Choose Excel File
        </label>
      </div>

      {showPreview && fileStream && (
        <ExcelPreview
          fileStream={fileStream}
          fileName={fileName}
          onClose={() => setShowPreview(false)}
        />
      )}
    </div>
  );
}

export default App; 