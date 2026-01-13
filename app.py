import React, { useState, useMemo } from 'react';
import { Upload, Download, CheckCircle, XCircle, AlertCircle, FileText, Search, Filter } from 'lucide-react';

const FAQCorrector = () => {
  const [file, setFile] = useState(null);
  const [data, setData] = useState([]);
  const [processedData, setProcessedData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterLevel, setFilterLevel] = useState('all');

  // Clean FAQ levels function
  const cleanFAQLevels = (text) => {
    if (!text || text === 'null' || text === 'undefined') {
      return [null, null, null, null, null];
    }

    text = String(text).replace(/"/g, '').trim();
    text = text.replace(/\n/g, '|');
    text = text.replace(/(?<=[a-z])(?=[A-Z])/g, ' | ');
    text = text.replace(/\s+/g, ' ');

    let parts = text.split('|').map(p => p.trim()).filter(p => p);
    parts = parts.slice(0, 5);

    while (parts.length < 5) {
      parts.push(null);
    }

    return parts;
  };

  // Generate question from row
  const generateQuestion = (level4, level5) => {
    if (level5 && level5 !== 'null') return level5;
    if (level4 && level4 !== 'null') return level4;
    return '';
  };

  // Handle file upload
  const handleFileUpload = async (e) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setLoading(true);

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = await import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs').then(m => m.read(data, { type: 'array' }));
        
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = await import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs').then(m => m.utils.sheet_to_json(worksheet));

        setData(jsonData);
        processData(jsonData);
        setLoading(false);
      };
      reader.readAsArrayBuffer(uploadedFile);
    } catch (error) {
      console.error('Error reading file:', error);
      setLoading(false);
    }
  };

  // Process data
  const processData = (rawData) => {
    const processed = rawData.map((row, idx) => {
      const faqColumn = row['Count of FAQ'] || row['FAQ'] || '';
      const [level1, level2, level3, level4, level5] = cleanFAQLevels(faqColumn);

      const faqCategory = level3 || '';
      const faqDescription = [level4, level5]
        .filter(l => l && l !== 'null')
        .join(' - ');
      
      const question = generateQuestion(level4, level5);

      return {
        id: idx + 1,
        originalFAQ: faqColumn,
        level1,
        level2,
        level3,
        level4,
        level5,
        faqCategory,
        faqDescription,
        question,
        ...row
      };
    });

    setProcessedData(processed);
  };

  // Download processed data
  const downloadProcessed = async () => {
    const XLSX = await import('https://cdn.sheetjs.com/xlsx-0.20.1/package/xlsx.mjs');
    
    const worksheet = XLSX.utils.json_to_sheet(processedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Processed FAQs');
    
    XLSX.writeFile(workbook, `FAQ_Corrected_${new Date().getTime()}.xlsx`);
  };

  // Filter and search
  const filteredData = useMemo(() => {
    return processedData.filter(row => {
      const matchesSearch = searchTerm === '' || 
        Object.values(row).some(val => 
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        );
      
      const matchesLevel = filterLevel === 'all' || 
        (filterLevel === 'complete' && row.level5) ||
        (filterLevel === 'incomplete' && !row.level5);
      
      return matchesSearch && matchesLevel;
    });
  }, [processedData, searchTerm, filterLevel]);

  // Statistics
  const stats = useMemo(() => {
    const total = processedData.length;
    const complete = processedData.filter(r => r.level5).length;
    const incomplete = total - complete;
    const withCategory = processedData.filter(r => r.faqCategory).length;

    return { total, complete, incomplete, withCategory };
  }, [processedData]);

  return (
    <div className="min-h-screen p-8" style={{
      background: 'linear-gradient(135deg, #FFF5E6 0%, #FFE4D6 100%)'
    }}>
      {/* Header */}
      <div className="max-w-7xl mx-auto mb-8">
        <h1 className="text-4xl font-bold mb-2" style={{ color: '#3D2C5C' }}>
          🔧 FAQ Corrector
        </h1>
        <p className="text-lg" style={{ color: '#574964' }}>
          Upload, process, and correct FAQ data with automatic level parsing
        </p>
      </div>

      {/* Instructions */}
      <div className="max-w-7xl mx-auto mb-6 bg-white rounded-2xl p-6 shadow-lg border-l-4" style={{ borderColor: '#574964' }}>
        <h3 className="font-semibold mb-3 flex items-center gap-2" style={{ color: '#574964' }}>
          <FileText size={20} />
          How to use:
        </h3>
        <ol className="space-y-2" style={{ color: '#2C2C2C' }}>
          <li><strong>1.</strong> Upload your Excel file containing FAQ data</li>
          <li><strong>2.</strong> The app will automatically parse and clean FAQ levels (1-5)</li>
          <li><strong>3.</strong> Review the processed data with search and filters</li>
          <li><strong>4.</strong> Download the corrected Excel file</li>
        </ol>
      </div>

      {/* Upload Section */}
      <div className="max-w-7xl mx-auto mb-6">
        <div className="bg-white rounded-2xl p-8 shadow-lg border-2 border-dashed" style={{ borderColor: '#9F8383' }}>
          <label className="flex flex-col items-center gap-4 cursor-pointer">
            <Upload size={48} style={{ color: '#574964' }} />
            <span className="text-lg font-semibold" style={{ color: '#574964' }}>
              {file ? file.name : 'Click to upload Excel file'}
            </span>
            <input
              type="file"
              accept=".xlsx,.xls"
              onChange={handleFileUpload}
              className="hidden"
            />
          </label>
        </div>
      </div>

      {loading && (
        <div className="max-w-7xl mx-auto mb-6 bg-white rounded-2xl p-6 shadow-lg text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 mx-auto mb-4" style={{ borderColor: '#574964' }}></div>
          <p style={{ color: '#574964' }}>Processing file...</p>
        </div>
      )}

      {/* Statistics */}
      {processedData.length > 0 && (
        <>
          <div className="max-w-7xl mx-auto mb-6 grid grid-cols-4 gap-4">
            <div className="bg-white rounded-xl p-6 shadow-lg border-l-4" style={{ borderColor: '#574964' }}>
              <div className="text-sm font-semibold mb-2" style={{ color: '#7B6B8E' }}>TOTAL RECORDS</div>
              <div className="text-3xl font-bold" style={{ color: '#3D2C5C' }}>{stats.total}</div>
            </div>
            <div className="bg-white rounded-xl p-6 shadow-lg border-l-4" style={{ borderColor: '#27AE60' }}>
              <div className="text-sm font-semibold mb-2" style={{ color: '#229954' }}>COMPLETE (5 LEVELS)</div>
              <div className="text-3xl font-bold" style={{ color: '#27AE60' }}>{stats.complete}</div>
            </div>
            <div className="bg-white rounded-xl p-6 shadow-lg border-l-4" style={{ borderColor: '#E74C3C' }}>
              <div className="text-sm font-semibold mb-2" style={{ color: '#C0392B' }}>INCOMPLETE</div>
              <div className="text-3xl font-bold" style={{ color: '#E74C3C' }}>{stats.incomplete}</div>
            </div>
            <div className="bg-white rounded-xl p-6 shadow-lg border-l-4" style={{ borderColor: '#3498DB' }}>
              <div className="text-sm font-semibold mb-2" style={{ color: '#2980B9' }}>WITH CATEGORY</div>
              <div className="text-3xl font-bold" style={{ color: '#3498DB' }}>{stats.withCategory}</div>
            </div>
          </div>

          {/* Search and Filter */}
          <div className="max-w-7xl mx-auto mb-6 bg-white rounded-2xl p-6 shadow-lg">
            <div className="flex gap-4">
              <div className="flex-1 relative">
                <Search className="absolute left-3 top-1/2 transform -translate-y-1/2" size={20} style={{ color: '#9E9E9E' }} />
                <input
                  type="text"
                  placeholder="Search in any field..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="w-full pl-10 pr-4 py-3 border-2 rounded-lg"
                  style={{ borderColor: '#E0E0E0', color: '#2C2C2C' }}
                />
              </div>
              <div className="relative">
                <Filter className="absolute left-3 top-1/2 transform -translate-y-1/2" size={20} style={{ color: '#9E9E9E' }} />
                <select
                  value={filterLevel}
                  onChange={(e) => setFilterLevel(e.target.value)}
                  className="pl-10 pr-8 py-3 border-2 rounded-lg appearance-none cursor-pointer"
                  style={{ borderColor: '#E0E0E0', color: '#2C2C2C' }}
                >
                  <option value="all">All Records</option>
                  <option value="complete">Complete (5 Levels)</option>
                  <option value="incomplete">Incomplete</option>
                </select>
              </div>
            </div>
          </div>

          {/* Download Button */}
          <div className="max-w-7xl mx-auto mb-6">
            <button
              onClick={downloadProcessed}
              className="w-full py-4 rounded-xl font-semibold text-white shadow-lg flex items-center justify-center gap-3 transition-all hover:shadow-xl hover:transform hover:-translate-y-1"
              style={{
                background: 'linear-gradient(135deg, #27AE60 0%, #229954 100%)'
              }}
            >
              <Download size={24} />
              Download Corrected Excel File ({filteredData.length} records)
            </button>
          </div>

          {/* Data Table */}
          <div className="max-w-7xl mx-auto bg-white rounded-2xl p-6 shadow-lg overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="border-b-2" style={{ borderColor: '#E0E0E0' }}>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>ID</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Level 1</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Level 2</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Level 3</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Level 4</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Level 5</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Category</th>
                    <th className="px-4 py-3 text-left text-sm font-semibold" style={{ color: '#574964' }}>Question</th>
                    <th className="px-4 py-3 text-center text-sm font-semibold" style={{ color: '#574964' }}>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredData.slice(0, 100).map((row, idx) => (
                    <tr key={row.id} className="border-b hover:bg-gray-50 transition-colors" style={{ borderColor: '#F5F5F5' }}>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.id}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.level1 || '-'}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.level2 || '-'}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.level3 || '-'}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.level4 || '-'}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.level5 || '-'}</td>
                      <td className="px-4 py-3 text-sm font-medium" style={{ color: '#574964' }}>{row.faqCategory || '-'}</td>
                      <td className="px-4 py-3 text-sm" style={{ color: '#2C2C2C' }}>{row.question || '-'}</td>
                      <td className="px-4 py-3 text-center">
                        {row.level5 ? (
                          <CheckCircle size={20} className="inline" style={{ color: '#27AE60' }} />
                        ) : (
                          <AlertCircle size={20} className="inline" style={{ color: '#E74C3C' }} />
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredData.length > 100 && (
                <div className="mt-4 text-center text-sm" style={{ color: '#7B6B8E' }}>
                  Showing first 100 of {filteredData.length} records. Download Excel for full data.
                </div>
              )}
            </div>
          </div>
        </>
      )}

      {/* Empty State */}
      {!loading && processedData.length === 0 && (
        <div className="max-w-7xl mx-auto bg-white rounded-2xl p-12 shadow-lg text-center">
          <FileText size={64} className="mx-auto mb-4" style={{ color: '#9E9E9E' }} />
          <p className="text-lg" style={{ color: '#7B6B8E' }}>
            Upload an Excel file to start processing FAQ data
          </p>
        </div>
      )}
    </div>
  );
};

export default FAQCorrector;
