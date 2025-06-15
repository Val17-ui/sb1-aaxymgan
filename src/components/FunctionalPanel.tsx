import React from 'react';
import { useDropzone } from 'react-dropzone';
import { FileUp, FileDown, Clock, Loader2, Cloud } from 'lucide-react';

interface FunctionalPanelProps {
  excelFile: File | null;
  templateFile: File | null;
  defaultDuration: number;
  processing: boolean;
  questions: any[];
  onExcelDrop: (files: File[]) => void;
  onTemplateDrop: (files: File[]) => void;
  onDefaultDurationChange: (duration: number) => void;
  onProcessExcel: () => void;
  onGeneratePPTX: () => void;
}

const FunctionalPanel: React.FC<FunctionalPanelProps> = ({
  excelFile,
  templateFile,
  defaultDuration,
  processing,
  questions,
  onExcelDrop,
  onTemplateDrop,
  onDefaultDurationChange,
  onProcessExcel,
  onGeneratePPTX,
}) => {
  const { getRootProps: getExcelRootProps, getInputProps: getExcelInputProps } = useDropzone({
    onDrop: onExcelDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    maxFiles: 1
  });

  const { getRootProps: getTemplateRootProps, getInputProps: getTemplateInputProps } = useDropzone({
    onDrop: onTemplateDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx']
    },
    maxFiles: 1
  });

  return (
    <div className="space-y-6">
      {/* Excel File Upload */}
      <div className="bg-white rounded-xl shadow-lg p-6">
        <h3 className="text-lg font-semibold text-gray-800 mb-4 flex items-center">
          <div className="w-2 h-5 bg-blue-500 rounded-full mr-3"></div>
          1. Upload Questions Excel
        </h3>
        <div 
          {...getExcelRootProps()} 
          className={`border-2 border-dashed rounded-lg p-8 flex flex-col items-center justify-center cursor-pointer transition-all duration-200 ${
            excelFile 
              ? 'border-green-400 bg-green-50 hover:bg-green-100' 
              : 'border-gray-300 hover:border-blue-400 hover:bg-blue-50'
          }`}
        >
          <input {...getExcelInputProps()} />
          <FileUp className={`w-12 h-12 mb-3 transition-colors ${
            excelFile ? 'text-green-500' : 'text-blue-500'
          }`} />
          {excelFile ? (
            <div className="text-center">
              <p className="text-sm font-medium text-gray-900 mb-1">{excelFile.name}</p>
              <p className="text-xs text-gray-500">
                {(excelFile.size / 1024).toFixed(2)} KB
              </p>
              <div className="mt-2 px-3 py-1 bg-green-100 text-green-800 text-xs rounded-full inline-block">
                ✓ Fichier chargé
              </div>
            </div>
          ) : (
            <div className="text-center">
              <p className="text-sm font-medium text-gray-700 mb-1">Glissez votre fichier Excel ici</p>
              <p className="text-xs text-gray-500">ou cliquez pour parcourir</p>
              <p className="text-xs text-gray-400 mt-2">Formats supportés: .xlsx, .xls</p>
            </div>
          )}
        </div>
        <button
          onClick={onProcessExcel}
          disabled={!excelFile || processing}
          className={`mt-4 w-full py-3 px-4 rounded-lg font-medium transition-all duration-200 ${
            !excelFile || processing
              ? 'bg-gray-200 text-gray-500 cursor-not-allowed'
              : 'bg-blue-600 text-white hover:bg-blue-700 hover:shadow-md'
          }`}
        >
          {processing ? (
            <span className="flex items-center justify-center">
              <Loader2 className="w-4 h-4 mr-2 animate-spin" />
              Traitement en cours...
            </span>
          ) : (
            'Traiter le fichier Excel'
          )}
        </button>
      </div>

      {/* Template PPTX Upload */}
      <div className="bg-white rounded-xl shadow-lg p-6">
        <h3 className="text-lg font-semibold text-gray-800 mb-4 flex items-center">
          <div className="w-2 h-5 bg-purple-500 rounded-full mr-3"></div>
          2. Upload Template PPTX (Optionnel)
        </h3>
        <div 
          {...getTemplateRootProps()} 
          className={`border-2 border-dashed rounded-lg p-8 flex flex-col items-center justify-center cursor-pointer transition-all duration-200 ${
            templateFile 
              ? 'border-green-400 bg-green-50 hover:bg-green-100' 
              : 'border-gray-300 hover:border-purple-400 hover:bg-purple-50'
          }`}
        >
          <input {...getTemplateInputProps()} />
          <FileUp className={`w-12 h-12 mb-3 transition-colors ${
            templateFile ? 'text-green-500' : 'text-purple-500'
          }`} />
          {templateFile ? (
            <div className="text-center">
              <p className="text-sm font-medium text-gray-900 mb-1">{templateFile.name}</p>
              <p className="text-xs text-gray-500">
                {(templateFile.size / 1024).toFixed(2)} KB
              </p>
              <div className="mt-2 px-3 py-1 bg-green-100 text-green-800 text-xs rounded-full inline-block">
                ✓ Template chargé
              </div>
            </div>
          ) : (
            <div className="text-center">
              <p className="text-sm font-medium text-gray-700 mb-1">Glissez votre template PPTX ici</p>
              <p className="text-xs text-gray-500">ou cliquez pour parcourir (optionnel)</p>
              <p className="text-xs text-gray-400 mt-2">Format supporté: .pptx</p>
            </div>
          )}
        </div>
      </div>

      {/* Cloud Images Info */}
      <div className="bg-blue-50 rounded-lg p-4 border border-blue-200">
        <div className="flex items-start">
          <Cloud className="w-5 h-5 text-blue-600 mr-3 flex-shrink-0 mt-0.5" />
          <div className="text-sm text-blue-800">
            <p className="font-medium mb-1">Images depuis le Cloud</p>
            <p className="text-blue-700">
              Votre fichier Excel peut inclure des URLs d'images depuis Google Drive, Dropbox, ou des liens directs. 
              Les images seront automatiquement téléchargées et intégrées dans le PowerPoint.
            </p>
          </div>
        </div>
      </div>

      {/* Default Duration Setting */}
      <div className="bg-white rounded-xl shadow-lg p-6">
        <h3 className="text-lg font-semibold text-gray-800 mb-4 flex items-center">
          <div className="w-2 h-5 bg-orange-500 rounded-full mr-3"></div>
          3. Durée par défaut des questions
        </h3>
        <div className="flex items-center space-x-3">
          <Clock className="w-5 h-5 text-gray-500" />
          <input
            type="number"
            min="5"
            max="120"
            value={defaultDuration}
            onChange={(e) => onDefaultDurationChange(parseInt(e.target.value) || 30)}
            className="w-24 px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-orange-500 focus:border-transparent"
          />
          <span className="text-gray-600">secondes</span>
        </div>
        <p className="text-xs text-gray-500 mt-2">
          Cette durée sera utilisée pour les questions qui n'ont pas de durée spécifiée dans l'Excel.
        </p>
      </div>

      {/* Generate Button */}
      <div className="bg-white rounded-xl shadow-lg p-6">
        <button
          onClick={onGeneratePPTX}
          disabled={questions.length === 0 || processing}
          className={`w-full py-4 px-6 rounded-lg font-semibold text-lg flex items-center justify-center transition-all duration-200 ${
            questions.length === 0 || processing
              ? 'bg-gray-200 text-gray-500 cursor-not-allowed'
              : 'bg-gradient-to-r from-indigo-600 to-purple-600 text-white hover:from-indigo-700 hover:to-purple-700 hover:shadow-lg transform hover:scale-[1.02]'
          }`}
        >
          {processing ? (
            <>
              <Loader2 className="w-5 h-5 mr-3 animate-spin" />
              Génération en cours...
            </>
          ) : (
            <>
              <FileDown className="w-5 h-5 mr-3" />
              Générer le PowerPoint OMBEA
            </>
          )}
        </button>
        {questions.length === 0 && (
          <p className="text-xs text-gray-500 text-center mt-2">
            Veuillez d'abord traiter un fichier Excel
          </p>
        )}
      </div>
    </div>
  );
};

export default FunctionalPanel;