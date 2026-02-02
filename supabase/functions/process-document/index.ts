import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import JSZip from "https://esm.sh/jszip@3.10.1";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

interface HighlightedSection {
  text: string;
  context: string;
  fullMatch: string;
}

interface ProcessedReplacement {
  originalText: string;
  newText: string;
  fullMatch: string;
}

function extractTextFromXml(xmlContent: string): string {
  // Extract all text from w:t tags
  const textMatches = xmlContent.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g);
  const texts: string[] = [];
  for (const match of textMatches) {
    texts.push(match[1]);
  }
  return texts.join(" ").trim();
}

function extractHighlightedSections(documentXml: string): HighlightedSection[] {
  const highlightedSections: HighlightedSection[] = [];
  
  // Pattern to find runs with highlight formatting
  // Looking for <w:r> elements that contain <w:highlight> in their <w:rPr> and have text in <w:t>
  const runPattern = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
  
  let contextBuffer: string[] = [];
  let runMatch: RegExpExecArray | null;
  
  while ((runMatch = runPattern.exec(documentXml)) !== null) {
    const runContent = runMatch[1];
    const fullMatch = runMatch[0];
    
    // Check if this run has highlight formatting
    const hasHighlight = /<w:highlight\b[^\/]*\/>/.test(runContent) || /<w:highlight\b[^>]*>/.test(runContent);
    
    // Extract text from this run
    const textMatches = runContent.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g);
    let runText = "";
    for (const tm of textMatches) {
      runText += tm[1];
    }
    
    if (runText.trim()) {
      contextBuffer.push(runText);
      if (contextBuffer.length > 10) {
        contextBuffer.shift();
      }
    }
    
    if (hasHighlight && runText.trim()) {
      highlightedSections.push({
        text: runText,
        context: contextBuffer.slice(-5).join(" "),
        fullMatch: fullMatch
      });
    }
  }
  
  return highlightedSections;
}

async function generateAllReplacements(
  instructionPrompt: string,
  sections: HighlightedSection[]
): Promise<Map<number, string>> {
  const apiKey = Deno.env.get("LOVABLE_API_KEY");
  
  if (!apiKey) {
    throw new Error("LOVABLE_API_KEY is not configured");
  }

  const maxRetries = 5;
  const baseDelay = 2000;

  // Build a single prompt with all sections
  const sectionsText = sections.map((s, i) => 
    `[Section ${i + 1}]\nContext: ${s.context}\nHighlighted text to replace: "${s.text}"`
  ).join("\n\n");

  for (let attempt = 0; attempt < maxRetries; attempt++) {
    const response = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "google/gemini-3-flash-preview",
        messages: [
          {
            role: "system",
            content: `You are a professional document editor. Replace highlighted placeholder text based on instructions.

RULES:
1. Generate ONLY replacement text - no explanations or formatting markers
2. Match tone and style of surrounding document
3. Use "Romit Acharya" for authorship/preparer/reviewer roles
4. Do NOT change client/beneficiary/third-party names
5. Generate reasonable professional estimates for numeric values
6. Keep similar length unless more detail is needed`
          },
          {
            role: "user",
            content: `INSTRUCTION PROMPT:\n${instructionPrompt}\n\nSECTIONS TO REPLACE:\n${sectionsText}\n\nProvide replacements for all sections.`
          }
        ],
        tools: [
          {
            type: "function",
            function: {
              name: "provide_replacements",
              description: "Provide replacement text for each highlighted section",
              parameters: {
                type: "object",
                properties: {
                  replacements: {
                    type: "array",
                    items: {
                      type: "object",
                      properties: {
                        section_number: { type: "number", description: "The section number (1-indexed)" },
                        replacement_text: { type: "string", description: "The replacement text" }
                      },
                      required: ["section_number", "replacement_text"]
                    }
                  }
                },
                required: ["replacements"]
              }
            }
          }
        ],
        tool_choice: { type: "function", function: { name: "provide_replacements" } },
        temperature: 0.7
      })
    });

    if (response.ok) {
      const data = await response.json();
      const toolCall = data.choices[0]?.message?.tool_calls?.[0];
      
      if (toolCall?.function?.arguments) {
        const parsed = JSON.parse(toolCall.function.arguments);
        const resultMap = new Map<number, string>();
        
        for (const r of parsed.replacements) {
          resultMap.set(r.section_number - 1, r.replacement_text); // Convert to 0-indexed
        }
        
        return resultMap;
      }
      throw new Error("Invalid AI response format");
    }

    if (response.status === 429) {
      if (attempt < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, attempt);
        console.log(`Rate limited, retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw new Error("Rate limited: The AI service is busy. Please try again in a few minutes.");
    }

    if (response.status === 402) {
      throw new Error("AI service credits exhausted. Please try again later.");
    }

    const errorText = await response.text();
    throw new Error(`AI API error: ${response.status} - ${errorText}`);
  }

  throw new Error("Max retries exceeded");
}

function replaceHighlightedText(
  documentXml: string,
  replacements: ProcessedReplacement[]
): string {
  let modifiedXml = documentXml;
  
  for (const replacement of replacements) {
    // Create the new run with the same structure but new text
    const newRun = replacement.fullMatch.replace(
      /<w:t([^>]*)>[^<]*<\/w:t>/g,
      `<w:t$1>${escapeXml(replacement.newText)}</w:t>`
    );
    
    modifiedXml = modifiedXml.replace(replacement.fullMatch, newRun);
  }
  
  return modifiedXml;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const formData = await req.formData();
    const instructionFile = formData.get("instructionPrompt");
    const clientDataFile = formData.get("clientData");

    console.log("Received instructionPrompt type:", typeof instructionFile, instructionFile?.constructor?.name);
    console.log("Received clientData type:", typeof clientDataFile, clientDataFile?.constructor?.name);

    if (!instructionFile || !clientDataFile) {
      return new Response(
        JSON.stringify({ error: "Both instruction prompt and client data files are required" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Handle file data - FormData returns File/Blob objects
    let instructionBuffer: ArrayBuffer;
    let clientBuffer: ArrayBuffer;

    // Type guard using duck typing for File/Blob
    const isFileOrBlob = (value: unknown): value is { arrayBuffer: () => Promise<ArrayBuffer> } => {
      return value !== null && typeof value === 'object' && 'arrayBuffer' in value;
    };

    if (isFileOrBlob(instructionFile)) {
      instructionBuffer = await instructionFile.arrayBuffer();
    } else {
      return new Response(
        JSON.stringify({ error: "Instruction prompt must be a file" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    if (isFileOrBlob(clientDataFile)) {
      clientBuffer = await clientDataFile.arrayBuffer();
    } else {
      return new Response(
        JSON.stringify({ error: "Client data must be a file" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    console.log("Instruction buffer size:", instructionBuffer.byteLength);
    console.log("Client buffer size:", clientBuffer.byteLength);

    // Verify the files start with PK (zip signature)
    const instructionView = new Uint8Array(instructionBuffer);
    const clientView = new Uint8Array(clientBuffer);
    
    console.log("Instruction file header:", instructionView[0], instructionView[1], instructionView[2], instructionView[3]);
    console.log("Client file header:", clientView[0], clientView[1], clientView[2], clientView[3]);

    if (instructionView[0] !== 0x50 || instructionView[1] !== 0x4B) {
      return new Response(
        JSON.stringify({ error: "Instruction file is not a valid .docx file (must be a ZIP archive)" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    if (clientView[0] !== 0x50 || clientView[1] !== 0x4B) {
      return new Response(
        JSON.stringify({ error: "Client data file is not a valid .docx file (must be a ZIP archive)" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Extract instruction prompt text from the docx
    const instructionZip = await JSZip.loadAsync(instructionBuffer);
    const instructionDocXml = instructionZip.file("word/document.xml");
    
    if (!instructionDocXml) {
      throw new Error("Invalid instruction prompt DOCX file");
    }
    
    const instructionXmlContent = await instructionDocXml.async("text");
    const instructionPrompt = extractTextFromXml(instructionXmlContent);

    if (!instructionPrompt) {
      return new Response(
        JSON.stringify({ error: "Could not extract text from instruction prompt file" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Process client data file
    const clientZip = await JSZip.loadAsync(clientBuffer);
    const clientDocXml = clientZip.file("word/document.xml");
    
    if (!clientDocXml) {
      throw new Error("Invalid client data DOCX file");
    }
    
    const documentXml = await clientDocXml.async("text");
    const highlightedSections = extractHighlightedSections(documentXml);

    if (highlightedSections.length === 0) {
      return new Response(
        JSON.stringify({ error: "No highlighted text found in the client data document" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Generate all replacements in a single batched API call
    console.log(`Processing ${highlightedSections.length} highlighted sections in a single batch request`);
    const replacementMap = await generateAllReplacements(instructionPrompt, highlightedSections);
    
    const replacements: ProcessedReplacement[] = highlightedSections.map((section, index) => ({
      originalText: section.text,
      newText: replacementMap.get(index) || section.text,
      fullMatch: section.fullMatch
    }));

    // Create the modified document
    const modifiedXml = replaceHighlightedText(documentXml, replacements);
    clientZip.file("word/document.xml", modifiedXml);
    
    const outputBuffer = await clientZip.generateAsync({ 
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 }
    });

    // Return the modified document
    return new Response(outputBuffer, {
      headers: {
        ...corsHeaders,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="updated-document.docx"`
      }
    });

  } catch (error: unknown) {
    console.error("Error processing document:", error);
    const errorMessage = error instanceof Error ? error.message : "Failed to process document";
    return new Response(
      JSON.stringify({ error: errorMessage }),
      { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
