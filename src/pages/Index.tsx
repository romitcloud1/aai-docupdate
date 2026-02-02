import { useState, useCallback } from "react";
import { FileUpload } from "@/components/FileUpload";
import { ProcessingStatus } from "@/components/ProcessingStatus";
import { Button } from "@/components/ui/button";
import { FileOutput, Sparkles } from "lucide-react";
import { supabase } from "@/integrations/supabase/client";

type ProcessingState = "idle" | "processing" | "success" | "error";

const Index = () => {
  const [instructionFile, setInstructionFile] = useState<File | null>(null);
  const [clientDataFile, setClientDataFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingState>("idle");
  const [statusMessage, setStatusMessage] = useState<string>("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadFilename, setDownloadFilename] = useState<string>("");

  const handleGenerate = useCallback(async () => {
    if (!instructionFile || !clientDataFile) return;

    setStatus("processing");
    setStatusMessage("Analyzing documents and generating content...");
    setDownloadUrl(null);

    try {
      const formData = new FormData();
      formData.append("instructionPrompt", instructionFile);
      formData.append("clientData", clientDataFile);

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
      let filename = "updated-document.docx";
      if (contentDisposition) {
        const match = contentDisposition.match(/filename="([^"]+)"/);
        if (match) {
          filename = match[1];
        }
      }

      // The response is a blob (the docx file)
      const blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      setDownloadUrl(url);
      setDownloadFilename(filename);
      setStatus("success");
      setStatusMessage("Document processed successfully! Click below to download.");
    } catch (err) {
      setStatus("error");
      setStatusMessage(err instanceof Error ? err.message : "An unexpected error occurred");
    }
  }, [instructionFile, clientDataFile]);

  const handleDownload = useCallback(() => {
    if (!downloadUrl) return;
    
    // Get filename from state or generate default
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = downloadFilename || "updated-document.docx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }, [downloadUrl, downloadFilename]);

  const isReady = instructionFile && clientDataFile;

  return (
    <div className="min-h-screen bg-background">
      <div className="max-w-2xl mx-auto px-4 py-12 sm:py-20">
        {/* Header */}
        <div className="text-center mb-12">
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-primary/10 mb-6">
            <FileOutput className="w-8 h-8 text-primary" />
          </div>
          <h1 className="text-3xl sm:text-4xl font-bold text-foreground mb-3">
            Document Processor
          </h1>
          <p className="text-muted-foreground max-w-md mx-auto">
            Upload your instruction prompt and client document to automatically update highlighted placeholders.
          </p>
        </div>

        {/* Upload Cards */}
        <div className="space-y-6 mb-8">
          <FileUpload
            label="Instruction Prompt"
            description="Word document (.docx) containing your generation instructions"
            onFileSelect={setInstructionFile}
            selectedFile={instructionFile}
          />

          <FileUpload
            label="Client Data"
            description="Word document (.docx) with highlighted text to replace"
            onFileSelect={setClientDataFile}
            selectedFile={clientDataFile}
          />
        </div>

        {/* Generate Button */}
        <Button
          onClick={handleGenerate}
          disabled={!isReady || status === "processing"}
          className="w-full h-12 text-base font-medium"
          size="lg"
        >
          <Sparkles className="w-5 h-5 mr-2" />
          Generate Updated Document
        </Button>

        {/* Status & Download */}
        <div className="mt-6 space-y-4">
          <ProcessingStatus status={status} message={statusMessage} />
          
          {downloadUrl && (
            <Button
              onClick={handleDownload}
              variant="outline"
              className="w-full h-12 text-base border-success text-success hover:bg-success/10"
            >
              <FileOutput className="w-5 h-5 mr-2" />
              Download Updated Document
            </Button>
          )}
        </div>

        {/* Footer */}
        <p className="text-center text-xs text-muted-foreground mt-12">
          Only highlighted text in your document will be modified. All other content remains unchanged.
        </p>
      </div>
    </div>
  );
};

export default Index;
