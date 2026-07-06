import * as React from "react";

export interface IDownloadButtonProps {
  fileUrl: string;
  fileName: string;
  className?: string;
}

const DownloadButton: React.FC<IDownloadButtonProps> = ({
  fileUrl,
  fileName,
  className,
}) => {
  const handleDownload = () => {
    const downloadUrl = `${
      window.location.origin
    }/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(fileUrl)}`;

    const link = document.createElement("a");

    link.href = downloadUrl;
    link.setAttribute("download", fileName);

    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <button onClick={handleDownload} className={className}>
      Download
    </button>
  );
};

export default DownloadButton;
