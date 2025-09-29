import React, { useState } from 'react';
import { Upload, Play, Download, CheckCircle, XCircle, AlertCircle, Loader } from 'lucide-react';

const StreamlitDataProcessor = () => {
  const [files, setFiles] = useState({
    crosses: null,
    parametric: null,
    packagePinout: null,
    recipe: null,
    lookup1: null,
    lookup2: null
  });
  const [processing, setProcessing] = useState(false);
  const [currentStep, setCurrentStep] = useState(0);
  const [results, setResults] = useState(null);
  const [loadLookup, setLoadLookup] = useState(true);

  const steps = [
    { name: 'Upload Files', icon: Upload },
    { name: 'Validate & Process', icon: Play },
    { name: 'Merge Data', icon: AlertCircle },
    { name: 'Generate Report', icon: CheckCircle },
    { name: 'Download Results', icon: Download }
  ];

  const fileInputs = [
    { key: 'crosses', label: 'Cross Reference File', required: true },
    { key: 'parametric', label: 'Parametric Data File', required: true },
    { key: 'packagePinout', label: 'Package & Pinout File', required: true },
    { key: 'recipe', label: 'Recipe File', required: true },
    { key: 'lookup1', label: 'Lookup Table 1', required: false },
    { key: 'lookup2', label: 'Lookup Table 2', required: false }
  ];

  const handleFileChange = (key, file) => {
    setFiles(prev => ({ ...prev, [key]: file }));
  };

  const handleProcess = async () => {
    setProcessing(true);
    setCurrentStep(1);
    
    // Simulate processing steps
    const stepDuration = 2000;
    
    for (let i = 1; i <= 4; i++) {
      await new Promise(resolve => setTimeout(resolve, stepDuration));
      setCurrentStep(i);
    }
    
    setResults({
      totalRecords: 156,
      crossMatches: 89,
      dropInA: 45,
      dropInB: 22,
      dropInC: 12,
      notDropIn: 10,
      differentPLs: 23,
      notFoundData: 44
    });
    
    setProcessing(false);
  };

  const canProcess = files.crosses && files.parametric && files.packagePinout && files.recipe;

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-8">
      <div className="max-w-6xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <h1 className="text-3xl font-bold text-gray-800 mb-2">
            Component Cross-Reference Validator
          </h1>
          <p className="text-gray-600">
            Upload your data files and validate component cross-references with automated grading
          </p>
        </div>

        {/* Progress Steps */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <div className="flex justify-between items-center">
            {steps.map((step, index) => {
              const Icon = step.icon;
              const isActive = index === currentStep;
              const isCompleted = index < currentStep;
              
              return (
                <div key={index} className="flex items-center">
                  <div className={`flex flex-col items-center ${index < steps.length - 1 ? 'mr-4' : ''}`}>
                    <div className={`w-12 h-12 rounded-full flex items-center justify-center transition-all ${
                      isCompleted ? 'bg-green-500 text-white' :
                      isActive ? 'bg-blue-500 text-white' :
                      'bg-gray-200 text-gray-400'
                    }`}>
                      <Icon size={24} />
                    </div>
                    <span className={`text-sm mt-2 font-medium ${
                      isActive ? 'text-blue-600' : 'text-gray-600'
                    }`}>
                      {step.name}
                    </span>
                  </div>
                  {index < steps.length - 1 && (
                    <div className={`h-1 w-16 mx-2 ${
                      index < currentStep ? 'bg-green-500' : 'bg-gray-200'
                    }`} />
                  )}
                </div>
              );
            })}
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          {/* File Upload Section */}
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-bold text-gray-800 mb-4 flex items-center">
              <Upload className="mr-2" size={24} />
              Upload Files
            </h2>
            
            <div className="space-y-4">
              {fileInputs.map(input => (
                <div key={input.key}>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    {input.label}
                    {input.required && <span className="text-red-500 ml-1">*</span>}
                  </label>
                  <div className="relative">
                    <input
                      type="file"
                      accept=".xlsx,.xls"
                      onChange={(e) => handleFileChange(input.key, e.target.files[0])}
                      className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                      disabled={processing}
                    />
                    {files[input.key] && (
                      <CheckCircle className="absolute right-3 top-3 text-green-500" size={20} />
                    )}
                  </div>
                </div>
              ))}

              <div className="flex items-center mt-4">
                <input
                  type="checkbox"
                  id="loadLookup"
                  checked={loadLookup}
                  onChange={(e) => setLoadLookup(e.target.checked)}
                  className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500"
                  disabled={processing}
                />
                <label htmlFor="loadLookup" className="ml-2 text-sm font-medium text-gray-700">
                  Load Lookup Tables for Normalization
                </label>
              </div>
            </div>

            <button
              onClick={handleProcess}
              disabled={!canProcess || processing}
              className={`w-full mt-6 py-3 px-4 rounded-lg font-semibold text-white transition-all ${
                canProcess && !processing
                  ? 'bg-blue-600 hover:bg-blue-700 active:scale-95'
                  : 'bg-gray-400 cursor-not-allowed'
              }`}
            >
              {processing ? (
                <span className="flex items-center justify-center">
                  <Loader className="animate-spin mr-2" size={20} />
                  Processing...
                </span>
              ) : (
                <span className="flex items-center justify-center">
                  <Play className="mr-2" size={20} />
                  Start Processing
                </span>
              )}
            </button>
          </div>

          {/* Results Section */}
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-bold text-gray-800 mb-4 flex items-center">
              <CheckCircle className="mr-2" size={24} />
              Processing Results
            </h2>
            
            {!results && !processing && (
              <div className="flex flex-col items-center justify-center h-64 text-gray-400">
                <AlertCircle size={64} className="mb-4" />
                <p>Upload files and start processing to see results</p>
              </div>
            )}

            {processing && (
              <div className="flex flex-col items-center justify-center h-64">
                <Loader className="animate-spin text-blue-500 mb-4" size={64} />
                <p className="text-gray-600 font-medium">Processing your data...</p>
                <p className="text-sm text-gray-500 mt-2">Step {currentStep} of {steps.length - 1}</p>
              </div>
            )}

            {results && !processing && (
              <div className="space-y-4">
                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-blue-50 p-4 rounded-lg">
                    <p className="text-sm text-gray-600">Total Records</p>
                    <p className="text-2xl font-bold text-blue-600">{results.totalRecords}</p>
                  </div>
                  <div className="bg-green-50 p-4 rounded-lg">
                    <p className="text-sm text-gray-600">Cross Matches</p>
                    <p className="text-2xl font-bold text-green-600">{results.crossMatches}</p>
                  </div>
                </div>

                <div className="border-t pt-4">
                  <h3 className="font-semibold text-gray-700 mb-3">Grade Distribution</h3>
                  <div className="space-y-2">
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Drop-in A</span>
                      <span className="font-semibold text-green-600">{results.dropInA}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Drop-in B</span>
                      <span className="font-semibold text-blue-600">{results.dropInB}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Drop-in C</span>
                      <span className="font-semibold text-yellow-600">{results.dropInC}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Not Drop-in</span>
                      <span className="font-semibold text-red-600">{results.notDropIn}</span>
                    </div>
                  </div>
                </div>

                <div className="border-t pt-4">
                  <h3 className="font-semibold text-gray-700 mb-3">Status Summary</h3>
                  <div className="space-y-2">
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Different PLs</span>
                      <span className="font-semibold">{results.differentPLs}</span>
                    </div>
                    <div className="flex justify-between items-center">
                      <span className="text-sm text-gray-600">Not Found Data</span>
                      <span className="font-semibold">{results.notFoundData}</span>
                    </div>
                  </div>
                </div>

                <button className="w-full mt-6 py-3 px-4 rounded-lg font-semibold text-white bg-green-600 hover:bg-green-700 transition-all active:scale-95 flex items-center justify-center">
                  <Download className="mr-2" size={20} />
                  Download Final Report
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Info Section */}
        <div className="bg-white rounded-lg shadow-lg p-6 mt-6">
          <h3 className="text-lg font-bold text-gray-800 mb-3">Processing Pipeline</h3>
          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 text-sm text-gray-600">
            <div>
              <h4 className="font-semibold text-gray-800 mb-2">Step 1: Validation</h4>
              <p>Compares PLc and PLx values, checks package/pinout similarity, and validates data availability.</p>
            </div>
            <div>
              <h4 className="font-semibold text-gray-800 mb-2">Step 2: Data Merging</h4>
              <p>Merges parametric data and package information for each part number from all sources.</p>
            </div>
            <div>
              <h4 className="font-semibold text-gray-800 mb-2">Step 3: Grading</h4>
              <p>Validates core features, tolerances, and upgrades to assign Drop-in grades (A, B, C, D).</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default StreamlitDataProcessor;
