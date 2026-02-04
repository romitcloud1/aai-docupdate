import { useState, useCallback } from "react";
import { FileUpload } from "@/components/FileUpload";
import { ProcessingStatus } from "@/components/ProcessingStatus";
import { ChangesViewer, FileChanges } from "@/components/ChangesViewer";
import { Button } from "@/components/ui/button";
import { FileOutput, Sparkles } from "lucide-react";
import advisoryAiLogo from "@/assets/advisoryai-logo.png";
import JSZip from "jszip";

type ProcessingState = "idle" | "processing" | "success" | "error";

const Index = () => {
  const [instructionFile, setInstructionFile] = useState<File | null>(null);
  const [clientDataFiles, setClientDataFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<ProcessingState>("idle");
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadFilename, setDownloadFilename] = useState<string>("");
  const [filesChanges, setFilesChanges] = useState<FileChanges[]>([]);

  const handleGenerate = useCallback(async () => {
    if (!instructionFile || clientDataFiles.length === 0) return;

    const totalFiles = clientDataFiles.length;
    setStatus("processing");
    setStatusMessage(totalFiles > 1 
      ? `Analyzing 1 of ${totalFiles} documents and generating content...`
      : `Analyzing document and generating content...`
    );
    setDownloadUrl(null);
    setFilesChanges([]);

    // Simulate progress updates for multi-file processing
    let progressInterval: NodeJS.Timeout | null = null;
    if (totalFiles > 1) {
      let currentDoc = 1;
      progressInterval = setInterval(() => {
        currentDoc = Math.min(currentDoc + 1, totalFiles);
        setStatusMessage(`Analyzing ${currentDoc} of ${totalFiles} documents and generating content...`);
        if (currentDoc >= totalFiles && progressInterval) {
          clearInterval(progressInterval);
        }
      }, 8000); // Update every 8 seconds (rough estimate per document)
    }

    try {
      const formData = new FormData();
      formData.append("instructionPrompt", instructionFile);
      
      // Append all client data files
      clientDataFiles.forEach((file) => {
        formData.append("clientData", file);
      });

      const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
      const supabaseKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY;

      const response = await fetch(`${supabaseUrl}/functions/v1/process-document`, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${supabaseKey}`,
          "apikey": supabaseKey,
        },
        body: formData,
      });

      if (progressInterval) clearInterval(progressInterval);

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to process document");
      }

      const data = await response.arrayBuffer();

      // Extract filename from Content-Disposition header
      const contentDisposition = response.headers.get("Content-Disposition");
      
      let filename = "processed-documents.zip";
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="([^"]+)"/);
        if (match) {
          filename = match[1];
        }
      }

      // Always a ZIP now - extract changes.json and create download URL
      const zip = await JSZip.loadAsync(data);
      
      // Extract changes metadata
      const changesFile = zip.file("_changes.json");
      console.log("Changes file found:", !!changesFile);
      if (changesFile) {
        const changesText = await changesFile.async("text");
        console.log("Changes content:", changesText.substring(0, 500));
        const changes: FileChanges[] = JSON.parse(changesText);
        console.log("Parsed changes:", changes.length, "files");
        setFilesChanges(changes);
      }

      // Create download blob (remove _changes.json for cleaner download)
      const downloadZip = new JSZip();
      const files = Object.keys(zip.files);
      for (const fileName of files) {
        if (fileName !== "_changes.json" && !zip.files[fileName].dir) {
          const fileData = await zip.files[fileName].async("arraybuffer");
          downloadZip.file(fileName, fileData);
        }
      }

      // If only one document file, extract it directly for download
      const docFiles = files.filter(f => f.endsWith('.docx'));
      if (docFiles.length === 1) {
        const docData = await zip.files[docFiles[0]].async("arraybuffer");
        const blob = new Blob([docData], { 
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
        });
        const url = URL.createObjectURL(blob);
        setDownloadUrl(url);
        setDownloadFilename(docFiles[0]);
      } else {
        // Multiple files - create clean ZIP without _changes.json
        const cleanZipBuffer = await downloadZip.generateAsync({ type: "arraybuffer" });
        const blob = new Blob([cleanZipBuffer], { type: "application/zip" });
        const url = URL.createObjectURL(blob);
        setDownloadUrl(url);
        setDownloadFilename(filename);
      }

      setStatus("success");
      setStatusMessage(`${totalFiles} document${totalFiles > 1 ? "s" : ""} processed successfully!`);
    } catch (err) {
      if (progressInterval) clearInterval(progressInterval);
      setStatus("error");
      setStatusMessage(err instanceof Error ? err.message : "An unexpected error occurred");
    }
  }, [instructionFile, clientDataFiles]);

  const handleDownload = useCallback(() => {
    if (!downloadUrl) return;
    
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = downloadFilename || "processed-documents.zip";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [downloadUrl, downloadFilename]);

  const isReady = instructionFile && clientDataFiles.length > 0;

  return (
    <div className="min-h-screen bg-background">
      <div className="max-w-2xl mx-auto px-4 py-12 sm:py-20">
        {/* Header */}
        <div className="text-center mb-12">
          <img 
            src={advisoryAiLogo} 
            alt="AdvisoryAI Logo" 
            className="h-12 mx-auto mb-6"
          />
          <h1 className="text-3xl sm:text-4xl font-bold text-foreground mb-3">
            Document Processor
          </h1>
          <p className="text-muted-foreground max-w-md mx-auto">
            Upload your instruction prompt and client documents to automatically update highlighted placeholders.
          </p>
        </div>

        {/* Upload Cards - Row Layout */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <FileUpload
            label="Instruction Prompt"
            description="Word document (.docx) containing your generation instructions"
            onFileSelect={setInstructionFile}
            selectedFile={instructionFile}
          />

          <FileUpload
            label="Client Data"
            description="Word documents (.docx) containing content to be processed"
            onFileSelect={() => {}}
            selectedFile={null}
            multiple={true}
            onFilesSelect={setClientDataFiles}
            selectedFiles={clientDataFiles}
          />
        </div>

        {/* Action Buttons - Centered Row */}
        <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
          <Button
            onClick={handleGenerate}
            disabled={!isReady || status === "processing"}
            className="w-full sm:w-auto px-8 h-12 text-base font-medium"
            size="lg"
          >
            <Sparkles className="w-5 h-5 mr-2" />
            Generate Updated Document{clientDataFiles.length > 1 ? "s" : ""}
          </Button>

          {downloadUrl && (
            <Button
              onClick={handleDownload}
              variant="outline"
              className="w-full sm:w-auto px-8 h-12 text-base border-success text-success hover:bg-success/10"
            >
              <FileOutput className="w-5 h-5 mr-2" />
              Download {clientDataFiles.length > 1 ? "ZIP" : "Document"}
            </Button>
          )}
        </div>

        {/* Status */}
        <div className="mt-6">
          <ProcessingStatus status={status} message={statusMessage} />
        </div>

        {/* Changes Viewer */}
        {filesChanges.length > 0 && (
          <div className="mt-6">
            <ChangesViewer filesChanges={filesChanges} />
          </div>
        )}
      </div>
    </div>
  );
};

export default Index;
