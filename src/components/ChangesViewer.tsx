import { useState } from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Eye, X } from "lucide-react";
import { ScrollArea } from "@/components/ui/scroll-area";

export interface FileChange {
  originalText: string;
  newText: string;
}

export interface FileChanges {
  fileName: string;
  changes: FileChange[];
}

interface ChangesViewerProps {
  filesChanges: FileChanges[];
}

export function ChangesViewer({ filesChanges }: ChangesViewerProps) {
  const [isOpen, setIsOpen] = useState(false);
  const [selectedFileIndex, setSelectedFileIndex] = useState(0);

  if (filesChanges.length === 0) return null;

  const hasMultipleFiles = filesChanges.length > 1;
  const currentFile = filesChanges[selectedFileIndex];

  return (
    <div className="w-full">
      {/* Toggle Button Row */}
      <div className="flex items-center justify-center gap-3">
        <Button
          variant="outline"
          onClick={() => setIsOpen(!isOpen)}
          className="gap-2"
        >
          {isOpen ? <X className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
          {isOpen ? "Hide Changes" : "View Changes"}
        </Button>

        {hasMultipleFiles && (
          <Select
            value={selectedFileIndex.toString()}
            onValueChange={(value) => setSelectedFileIndex(parseInt(value))}
          >
            <SelectTrigger className="w-[220px] bg-background">
              <SelectValue placeholder="Select file" />
            </SelectTrigger>
            <SelectContent className="bg-background z-50">
              {filesChanges.map((file, index) => (
                <SelectItem key={index} value={index.toString()}>
                  {file.fileName}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        )}
      </div>

      {/* Changes Panel */}
      {isOpen && currentFile && (
        <Card className="mt-4 border">
          <CardHeader className="pb-3">
            <CardTitle className="text-lg">
              Changes in {currentFile.fileName}
              <span className="text-muted-foreground text-sm font-normal ml-2">
                ({currentFile.changes.length} change{currentFile.changes.length !== 1 ? "s" : ""})
              </span>
            </CardTitle>
          </CardHeader>
          <CardContent>
            {currentFile.changes.length === 0 ? (
              <p className="text-muted-foreground text-center py-4">
                No changes were made to this file.
              </p>
            ) : (
              <ScrollArea className="h-[400px] pr-4">
                <div className="space-y-4">
                  {currentFile.changes.map((change, idx) => (
                    <div key={idx} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      {/* Original - Red */}
                      <div className="p-3 rounded-md bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-900">
                        <div className="text-xs font-medium text-red-600 dark:text-red-400 mb-1 uppercase tracking-wide">
                          Original
                        </div>
                        <p className="text-sm text-red-800 dark:text-red-200 break-words">
                          {change.originalText}
                        </p>
                      </div>

                      {/* New - Green */}
                      <div className="p-3 rounded-md bg-green-50 dark:bg-green-950/30 border border-green-200 dark:border-green-900">
                        <div className="text-xs font-medium text-green-600 dark:text-green-400 mb-1 uppercase tracking-wide">
                          New
                        </div>
                        <p className="text-sm text-green-800 dark:text-green-200 break-words">
                          {change.newText}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
              </ScrollArea>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}
