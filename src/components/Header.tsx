import React from 'react';
import { PresentationIcon } from 'lucide-react';

const Header: React.FC = () => {
  return (
    <header className="bg-indigo-700 text-white shadow-md">
      <div className="container mx-auto px-4 py-5 flex items-center justify-between">
        <div className="flex items-center space-x-3">
          <PresentationIcon className="h-8 w-8" />
          <div>
            <h1 className="text-2xl font-bold">OMBEA PowerPoint Generator</h1>
            <p className="text-indigo-200 text-sm">Create interactive voting slides automatically</p>
          </div>
        </div>
        <div className="hidden md:flex items-center space-x-4">
          <span className="px-3 py-1 bg-indigo-800 rounded-full text-xs font-medium">
            v1.0.0
          </span>
        </div>
      </div>
    </header>
  );
};

export default Header;