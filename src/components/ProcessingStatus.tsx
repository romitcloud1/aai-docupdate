import { Loader2, CheckCircle2, AlertCircle } from "lucide-react";

type Status = "idle" | "processing" | "success" | "error";

interface ProcessingStatusProps {
  status: Status;
  message?: string;
}

export function ProcessingStatus({ status, message }: ProcessingStatusProps) {
  if (status === "idle") return null;

  return (
    <div className="flex items-center gap-3 p-4 rounded-xl bg-card border border-border">
      {status === "processing" && (
        <>
          <Loader2 className="w-5 h-5 text-primary animate-spin" />
          <span className="text-sm text-foreground">
            {message || "Processing your documents..."}
          </span>
        </>
      )}
      {status === "success" && (
        <>
          <CheckCircle2 className="w-5 h-5 text-success" />
          <span className="text-sm text-foreground">
            {message || "Document processed successfully!"}
          </span>
        </>
      )}
      {status === "error" && (
        <>
          <AlertCircle className="w-5 h-5 text-destructive" />
          <span className="text-sm text-destructive">
            {message || "An error occurred while processing."}
          </span>
        </>
      )}
    </div>
  );
}
