import React from 'react';
import { Info } from 'lucide-react';
import { ConfigOptions } from '../../types';

// Local ConfigOptions interface removed, now imported from ../../types

interface ConfigPanelProps {
  config: ConfigOptions;
  onChange: (config: ConfigOptions) => void;
}

interface TooltipProps {
  text: string;
  children: React.ReactNode;
}

const Tooltip: React.FC<TooltipProps> = ({ text, children }) => {
  return (
    <div className="group relative inline-block">
      {children}
      <div className="invisible group-hover:visible absolute z-10 w-64 p-2 mt-1 text-xs text-white bg-gray-900 rounded-md shadow-lg opacity-0 group-hover:opacity-100 transition-opacity duration-200 left-0">
        {text}
        <div className="absolute -top-1 left-3 w-2 h-2 bg-gray-900 rotate-45"></div>
      </div>
    </div>
  );
};

const ConfigPanel: React.FC<ConfigPanelProps> = ({ config, onChange }) => {
  const updateConfig = (key: keyof ConfigOptions, value: any) => {
    onChange({ ...config, [key]: value });
  };

  const answerStyleOptions = [
    { value: 'ppBulletAlphaUCParenRight', label: 'A)-J)' },
    { value: 'ppBulletAlphaUCPeriod', label: 'A.-J.' },
    { value: 'ppBulletArabicParenRight', label: '1)-10)' },
    { value: 'ppBulletArabicPeriod', label: '1.-10.' },
  ];

  const formatOptions = [
    { value: 'Response_Count', label: 'Nombre' },
    { value: 'Percentage_Integer', label: '0%' },
    { value: 'Percentage_One_Decimal', label: '0,0%' },
    { value: 'Percentage_Two_Decimal', label: '0,00%' },
    { value: 'Percentage_Three_Decimal', label: '0,000%' },
  ];

  const timeLimitOptions = [
    { value: 0, label: 'Pas de limite' },
    { value: 10, label: '10 secondes' },
    { value: 20, label: '20 secondes' },
    { value: 30, label: '30 secondes' },
    { value: 40, label: '40 secondes' },
    { value: 50, label: '50 secondes' },
  ];

  return (
    <div className="bg-white rounded-xl shadow-lg p-6 h-fit">
      <h2 className="text-xl font-bold text-gray-800 mb-6 flex items-center">
        <div className="w-2 h-6 bg-indigo-600 rounded-full mr-3"></div>
        Configuration OMBEA
      </h2>
      
      <div className="space-y-6">
        {/* Ouverture du vote */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Ouverture du vote
            </label>
            <Tooltip text="Détermine si le vote commence dès l'affichage de la diapositive ou s'il faut l'activer manuellement.">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.pollStartMode}
            onChange={(e) => updateConfig('pollStartMode', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            <option value="Automatic">Automatique</option>
            <option value="Manual">Manuel</option>
          </select>
        </div>

        {/* Style des réponses */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Style des réponses
            </label>
            <Tooltip text="Définit le format d'affichage des options de réponse (lettres majuscules/minuscules, parenthèses/points).">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.answersBulletStyle}
            onChange={(e) => updateConfig('answersBulletStyle', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            {answerStyleOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        {/* Format de réponse */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Format de réponse
            </label>
            <Tooltip text="Choisit comment afficher les résultats du vote : en nombre de réponses ou en pourcentage avec différents niveaux de précision.">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.chartValueLabelFormat}
            onChange={(e) => updateConfig('chartValueLabelFormat', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            {formatOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        {/* Durée du compte à rebours */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Durée du compte à rebours
            </label>
            <Tooltip text="Temps limite en secondes pour répondre au vote. 0 = pas de limite de temps.">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.pollTimeLimit}
            onChange={(e) => updateConfig('pollTimeLimit', parseInt(e.target.value))}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            {timeLimitOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        {/* Déclenchement du chrono */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Déclenchement du chrono
            </label>
            <Tooltip text="Détermine si le compte à rebours démarre automatiquement ou doit être activé manuellement.">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.pollCountdownStartMode}
            onChange={(e) => updateConfig('pollCountdownStartMode', e.target.value)}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            <option value="Automatic">Automatique</option>
            <option value="Manual">Manuel</option>
          </select>
        </div>

        {/* Réponses par participant */}
        <div className="space-y-2">
          <div className="flex items-center space-x-2">
            <label className="text-sm font-medium text-gray-700">
              Réponses par participant
            </label>
            <Tooltip text="Nombre maximum de réponses qu'un participant peut sélectionner pour cette question.">
              <Info className="w-4 h-4 text-gray-400 hover:text-gray-600 cursor-help" />
            </Tooltip>
          </div>
          <select
            value={config.pollMultipleResponse}
            onChange={(e) => updateConfig('pollMultipleResponse', parseInt(e.target.value))}
            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 focus:border-transparent bg-white"
          >
            {Array.from({ length: 10 }, (_, i) => i + 1).map((num) => (
              <option key={num} value={num}>
                {num}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* Configuration Summary */}
      <div className="mt-8 p-4 bg-indigo-50 rounded-lg border border-indigo-200">
        <h3 className="text-sm font-medium text-indigo-800 mb-2">Résumé de la configuration</h3>
        <div className="text-xs text-indigo-700 space-y-1">
          <div>Vote: {config.pollStartMode === 'Automatic' ? 'Automatique' : 'Manuel'}</div>
          <div>Style: {answerStyleOptions.find(opt => opt.value === config.answersBulletStyle)?.label}</div>
          <div>Format: {formatOptions.find(opt => opt.value === config.chartValueLabelFormat)?.label}</div>
          <div>Durée: {timeLimitOptions.find(opt => opt.value === config.pollTimeLimit)?.label}</div>
          <div>Chrono: {config.pollCountdownStartMode === 'Automatic' ? 'Auto' : 'Manuel'}</div>
          <div>Réponses max: {config.pollMultipleResponse}</div>
        </div>
      </div>
    </div>
  );
};

export default ConfigPanel;