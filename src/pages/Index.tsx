import { useState, useCallback } from "react";
import { FileUpload } from "@/components/FileUpload";
import { ProcessingStatus } from "@/components/ProcessingStatus";
import { Button } from "@/components/ui/button";
import { FileOutput, Sparkles } from "lucide-react";
import advisoryAiLogo from "@/assets/advisoryai-logo.png";

type ProcessingState = "idle" | "processing" | "success" | "error";

const Index = () => {
  const [instructionFile, setInstructionFile] = useState<File | null>(null);
  const [clientDataFiles, setClientDataFiles] = useState<File[]>([]);
  const [status, setStatus] = useState<ProcessingState>("idle");
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadFilename, setDownloadFilename] = useState<string>("");

  const handleGenerate = useCallback(async () => {
    if (!instructionFile || clientDataFiles.length === 0) return;

    setStatus("processing");
    setStatusMessage(`Analyzing ${clientDataFiles.length} document${clientDataFiles.length > 1 ? "s" : ""} and generating content...`);
    setDownloadUrl(null);

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

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || "Failed to process document");
      }

      const data = await response.arrayBuffer();

      // Extract filename from Content-Disposition header
      const contentDisposition = response.headers.get("Content-Disposition");
      const contentType = response.headers.get("Content-Type");
      
      let filename = clientDataFiles.length > 1 ? "processed-documents.zip" : "updated-document.docx";
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="([^"]+)"/);
        if (match) {
          filename = match[1];
        }
      }

      // Determine blob type based on response
      const blobType = contentType || (clientDataFiles.length > 1 
        ? "application/zip" 
        : "application/vnd.openxmlformats-officedocument.wordprocessingml.document");

      const blob = new Blob([data], { type: blobType });
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadFilename(filename);
      setStatus("success");
      setStatusMessage(`${clientDataFiles.length} document${clientDataFiles.length > 1 ? "s" : ""} processed successfully! Click below to download.`);
    } catch (err) {
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
      </div>
    </div>
  );
};

export default Index;
