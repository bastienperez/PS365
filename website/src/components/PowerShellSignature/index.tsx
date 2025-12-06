import React from 'react';

interface PowerShellSignatureProps {
  signature: string;
}

export default function PowerShellSignature({ signature }: PowerShellSignatureProps) {
  // Parse the PowerShell signature and apply colors
  const formatSignature = (sig: string) => {
    return sig
      // Function names (start of line until first space or bracket)
      .replace(/^([A-Za-z-]+)/g, '<span class="ps-function">$1</span>')
      // Parameters (things that start with -)
      .replace(/(-[A-Za-z]+)/g, '<span class="ps-parameter">$1</span>')
      // Types (things between < >)
      .replace(/(<[^>]+>)/g, '<span class="ps-type">$1</span>')
      // Brackets and syntax characters
      .replace(/(\[|\]|\|)/g, '<span class="ps-syntax-char">$1</span>');
  };

  return (
    <div 
      className="ps-syntax"
      dangerouslySetInnerHTML={{ __html: formatSignature(signature) }}
    />
  );
}