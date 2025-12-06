import React from 'react';

interface PowerShellSignatureProps {
  command: string;
  parameters?: Array<{
    name: string;
    type: string;
    required?: boolean;
    description?: string;
  }>;
}

export default function PowerShellSignature({ command, parameters = [] }: PowerShellSignatureProps) {
  return (
    <div className="command-signature">
      <div style={{ marginBottom: '0.5rem' }}>
        <span className="command-name">{command}</span>
        {parameters.map((param, index) => (
          <span key={index}>
            {' '}
            {!param.required && <span className="optional">[</span>}
            <span className="parameter">[{param.name}]</span>
            {' '}
            <span className="parameter-type">&lt;{param.type}&gt;</span>
            {!param.required && <span className="optional">]</span>}
          </span>
        ))}
        <span className="optional"> [&lt;CommonParameters&gt;]</span>
      </div>
    </div>
  );
}