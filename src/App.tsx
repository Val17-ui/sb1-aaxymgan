import { useState, useCallback } from 'react';
import { CheckCircle, AlertCircle, Loader2 } from 'lucide-react';
import { processExcel } from './utils/excelProcessor';
import { generatePPTX } from './utils/pptxGenerator';
import Header from './components/Header';
import ConfigPanel, { ConfigOptions } from './components/ConfigPanel';
import FunctionalPanel from './components/FunctionalPanel';

function App() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState<boolean>(false);
  const [status, setStatus] = useState<{ message: string; type: 'info' | 'success' | 'error' } | null>(null);
  const [questions, setQuestions] = useState<any[]>([]);
  
  // Configuration OMBEA avec valeurs par défaut
  const [config, setConfig] = useState<ConfigOptions>({
    pollStartMode: 'Automatic',
    answersBulletStyle: 'ppBulletAlphaUCParenRight',
    chartValueLabelFormat: 'Response_Count',
    pollTimeLimit: 0,
    pollCountdownStartMode: 'Automatic',
    pollMultipleResponse: 1,
  });

  const onExcelDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      setExcelFile(acceptedFiles[0]);
      setStatus({ message: 'Fichier Excel chargé. Prêt à traiter.', type: 'info' });
    }
  }, []);

  const onTemplateDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      setTemplateFile(acceptedFiles[0]);
      setStatus({ message: 'Template PPTX chargé (optionnel).', type: 'info' });
    }
  }, []);

  const handleProcessExcel = async () => {
    if (!excelFile) {
      setStatus({ message: 'Veuillez d\'abord charger un fichier Excel.', type: 'error' });
      return;
    }

    setProcessing(true);
    setStatus({ message: 'Traitement du fichier Excel...', type: 'info' });

    try {
      const result = await processExcel(excelFile);
      setQuestions(result);
      
      const questionsWithImages = result.filter(q => q.imageUrl).length;
      const message = questionsWithImages > 0
        ? `${result.length} questions traitées avec succès (${questionsWithImages} avec images).`
        : `${result.length} questions traitées avec succès.`;
      
      setStatus({ message, type: 'success' });
    } catch (error) {
      console.error('Erreur lors du traitement Excel:', error);
      setStatus({ 
        message: `Erreur lors du traitement Excel: ${error instanceof Error ? error.message : 'Erreur inconnue'}`, 
        type: 'error' 
      });
    } finally {
      setProcessing(false);
    }
  };

  const handleGeneratePPTX = async () => {
    if (questions.length === 0) {
      setStatus({ message: 'Veuillez d\'abord traiter le fichier Excel.', type: 'error' });
      return;
    }
  
    setProcessing(true);
    setStatus({ message: 'Génération du fichier PPTX...', type: 'info' });
  
    try {
      // Passer la configuration OMBEA à la fonction de génération
      await generatePPTX(templateFile, questions, { 
        ombeaConfig: config 
      });
      setStatus({ message: 'Fichier PPTX généré avec succès !', type: 'success' });
    } catch (error) {
      console.error('Erreur lors de la génération PPTX:', error);
      setStatus({ 
        message: `Erreur lors de la génération PPTX: ${error instanceof Error ? error.message : 'Erreur inconnue'}`, 
        type: 'error' 
      });
    } finally {
      setProcessing(false);
    }
  };

  // Fonction pour vérifier si une URL d'image est valide
  const isValidImageUrl = (url: string | undefined): boolean => {
    if (!url) return false;
    try {
      const urlObj = new URL(url);
      return ['http:', 'https:', 'data:'].includes(urlObj.protocol);
    } catch {
      return false;
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-100">
      <Header />
      
      <main className="container mx-auto px-4 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Configuration Panel - Left Side */}
          <div className="lg:col-span-1">
            <ConfigPanel config={config} onChange={setConfig} />
          </div>

          {/* Functional Panel - Right Side */}
          <div className="lg:col-span-2">
            <FunctionalPanel
              excelFile={excelFile}
              templateFile={templateFile}
              defaultDuration={30} // Valeur par défaut pour compatibilité
              processing={processing}
              questions={questions}
              onExcelDrop={onExcelDrop}
              onTemplateDrop={onTemplateDrop}
              onDefaultDurationChange={() => {}} // Fonction vide pour compatibilité
              onProcessExcel={handleProcessExcel}
              onGeneratePPTX={handleGeneratePPTX}
            />
          </div>
        </div>

        {/* Status Message */}
        {status && (
          <div className={`mt-8 rounded-xl p-4 flex items-start shadow-lg ${
            status.type === 'success' ? 'bg-green-100 text-green-800 border border-green-200' :
            status.type === 'error' ? 'bg-red-100 text-red-800 border border-red-200' :
            'bg-blue-100 text-blue-800 border border-blue-200'
          }`}>
            {status.type === 'success' ? (
              <CheckCircle className="w-5 h-5 mr-3 flex-shrink-0 mt-0.5" />
            ) : status.type === 'error' ? (
              <AlertCircle className="w-5 h-5 mr-3 flex-shrink-0 mt-0.5" />
            ) : (
              <Loader2 className="w-5 h-5 mr-3 flex-shrink-0 mt-0.5 animate-spin" />
            )}
            <p className="font-medium">{status.message}</p>
          </div>
        )}

        {/* Preview Section */}
        {questions.length > 0 && (
          <div className="mt-8 bg-white rounded-xl shadow-lg p-6">
            <h2 className="text-xl font-bold text-gray-800 mb-6 flex items-center">
              <div className="w-2 h-6 bg-green-500 rounded-full mr-3"></div>
              Aperçu des Questions ({questions.length})
            </h2>
            <div className="overflow-x-auto">
              <table className="min-w-full divide-y divide-gray-200">
                <thead className="bg-gray-50">
                  <tr>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      #
                    </th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Question
                    </th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Réponse Correcte
                    </th>
                    <th scope="col" className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                      Image
                    </th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {questions.map((q, idx) => {
                    const hasValidImageUrl = isValidImageUrl(q.imageUrl);
                    
                    return (
                      <tr key={idx} className={idx % 2 === 0 ? 'bg-white' : 'bg-gray-50'}>
                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">
                          {idx + 1}
                        </td>
                        <td className="px-6 py-4 text-sm text-gray-500 max-w-xs">
                          <div className="truncate" title={q.question}>
                            {q.question.length > 60 ? q.question.substring(0, 60) + '...' : q.question}
                          </div>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm">
                          <span className={`px-2 py-1 text-xs font-medium rounded-full ${
                            q.correctAnswer 
                              ? 'bg-green-100 text-green-800' 
                              : 'bg-red-100 text-red-800'
                          }`}>
                            {q.correctAnswer ? 'Vrai' : 'Faux'}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-sm">
                          {q.imageUrl ? (
                            <div className="flex items-center">
                              {hasValidImageUrl ? (
                                <>
                                  <div className="w-2 h-2 bg-green-500 rounded-full mr-2"></div>
                                  <span className="text-green-600 text-xs truncate max-w-[150px]" title={q.imageUrl}>
                                    {new URL(q.imageUrl).hostname}
                                  </span>
                                </>
                              ) : (
                                <>
                                  <div className="w-2 h-2 bg-yellow-500 rounded-full mr-2"></div>
                                  <span className="text-yellow-600 text-xs">URL invalide</span>
                                </>
                              )}
                            </div>
                          ) : (
                            <span className="text-gray-400">-</span>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            
            {/* Info about supported image sources */}
            {questions.some(q => q.imageUrl) && (
              <div className="mt-6 p-4 bg-gray-50 border border-gray-200 rounded-lg">
                <p className="text-sm text-gray-700 font-medium mb-2">Sources d'images supportées :</p>
                <div className="grid grid-cols-2 gap-2 text-sm text-gray-600">
                  <div>• Google Drive (lien de partage)</div>
                  <div>• Dropbox (lien de partage)</div>
                  <div>• URLs d'images directes (https://...)</div>
                  <div>• OneDrive (lien direct)</div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

export default App;