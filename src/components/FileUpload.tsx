import { useCallback, useState } from "react";
import { Upload, FileText, X } from "lucide-react";
import { cn } from "@/lib/utils";

interface FileUploadProps {
  label: string;
  description: string;
  accept?: string;
  onFileSelect: (file: File | null) => void;
  selectedFile: File | null;
  multiple?: boolean;
  onFilesSelect?: (files: File[]) => void;
  selectedFiles?: File[];
}

export function FileUpload({
  label,
  description,
  accept = ".docx",
  onFileSelect,
  selectedFile,
  multiple = false,
  onFilesSelect,
  selectedFiles = [],
}: FileUploadProps) {
  const [isDragging, setIsDragging] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent<HTMLDivElement>) => {
      e.preventDefault();
      setIsDragging(false);
      
      if (multiple && onFilesSelect) {
        const files = Array.from(e.dataTransfer.files).filter(f => f.name.endsWith(".docx"));
        if (files.length > 0) {
          onFilesSelect([...selectedFiles, ...files]);
        }
      } else {
        const file = e.dataTransfer.files[0];
        if (file && file.name.endsWith(".docx")) {
          onFileSelect(file);
        }
      }
    },
    [onFileSelect, multiple, onFilesSelect, selectedFiles]
  );

  const handleDragOver = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (multiple && onFilesSelect) {
        const files = Array.from(e.target.files || []);
        if (files.length > 0) {
          onFilesSelect([...selectedFiles, ...files]);
        }
      } else {
        const file = e.target.files?.[0];
        if (file) {
          onFileSelect(file);
        }
      }
      // Reset input so same file can be selected again
      e.target.value = "";
    },
    [onFileSelect, multiple, onFilesSelect, selectedFiles]
  );

  const handleRemove = useCallback(
    (e: React.MouseEvent) => {
      e.stopPropagation();
      onFileSelect(null);
    },
    [onFileSelect]
  );

  const handleRemoveFile = useCallback(
    (e: React.MouseEvent, index: number) => {
      e.stopPropagation();
      if (onFilesSelect) {
        const newFiles = selectedFiles.filter((_, i) => i !== index);
        onFilesSelect(newFiles);
      }
    },
    [onFilesSelect, selectedFiles]
  );

  const hasFiles = multiple ? selectedFiles.length > 0 : !!selectedFile;

  return (
    <div className="space-y-2">
      <label className="text-sm font-medium text-foreground">{label}</label>
      <div
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        className={cn(
          "relative border-2 border-dashed rounded-xl p-8 transition-all duration-200 cursor-pointer",
          "hover:border-primary/50 hover:bg-primary/5",
          isDragging && "border-primary bg-primary/10",
          hasFiles ? "border-success bg-success/5" : "border-border"
        )}
      >
        <input
          type="file"
          accept={accept}
          multiple={multiple}
          onChange={handleFileChange}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        />
        <div className="flex flex-col items-center gap-3 text-center pointer-events-none">
          {multiple && selectedFiles.length > 0 ? (
            <>
              <div className="w-12 h-12 rounded-full bg-success/10 flex items-center justify-center">
                <FileText className="w-6 h-6 text-success" />
              </div>
              <div className="space-y-2 w-full max-w-md">
                <p className="text-sm font-medium text-foreground">
                  {selectedFiles.length} file{selectedFiles.length > 1 ? "s" : ""} selected
                </p>
                <div className="space-y-1">
                  {selectedFiles.map((file, index) => (
                    <div
                      key={`${file.name}-${index}`}
                      className="flex items-center justify-between bg-background/50 rounded px-3 py-1.5 text-xs"
                    >
                      <span className="truncate max-w-[200px]">{file.name}</span>
                      <button
                        onClick={(e) => handleRemoveFile(e, index)}
                        className="pointer-events-auto ml-2 text-destructive hover:text-destructive/80"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </div>
                  ))}
                </div>
                <p className="text-xs text-muted-foreground">
                  Click or drag to add more files
                </p>
              </div>
            </>
          ) : selectedFile ? (
            <>
              <div className="w-12 h-12 rounded-full bg-success/10 flex items-center justify-center">
                <FileText className="w-6 h-6 text-success" />
              </div>
              <div className="space-y-1">
                <p className="text-sm font-medium text-foreground">{selectedFile.name}</p>
                <p className="text-xs text-muted-foreground">
                  {(selectedFile.size / 1024).toFixed(1)} KB
                </p>
              </div>
              <button
                onClick={handleRemove}
                className="pointer-events-auto flex items-center gap-1 text-xs text-destructive hover:underline"
              >
                <X className="w-3 h-3" />
                Remove
              </button>
            </>
          ) : (
            <>
              <div className="w-12 h-12 rounded-full bg-muted flex items-center justify-center">
                <Upload className="w-6 h-6 text-muted-foreground" />
              </div>
              <div className="space-y-1">
                <p className="text-sm text-foreground">
                  <span className="font-medium text-primary">Click to upload</span> or drag and drop
                </p>
                <p className="text-xs text-muted-foreground">{description}</p>
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
}
