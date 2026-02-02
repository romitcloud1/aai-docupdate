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

async function fetchMarketData(): Promise<string> {
  const currentDate = new Date();
  const formattedDate = currentDate.toLocaleDateString('en-GB', { 
    day: 'numeric', 
    month: 'long', 
    year: 'numeric' 
  });

  try {
    // Fetch UK FTSE 100 data from Yahoo Finance (free, no API key needed)
    const ftseResponse = await fetch(
      'https://query1.finance.yahoo.com/v8/finance/chart/%5EFTSE?interval=1d&range=1d',
      { headers: { 'User-Agent': 'Mozilla/5.0' } }
    );
    
    let ftseData = "";
    if (ftseResponse.ok) {
      const ftseJson = await ftseResponse.json();
      const result = ftseJson?.chart?.result?.[0];
      if (result) {
        const price = result.meta?.regularMarketPrice;
        const prevClose = result.meta?.previousClose;
        if (price && prevClose) {
          const changePercent = ((price - prevClose) / prevClose * 100).toFixed(2);
          ftseData = `FTSE 100: ${price.toFixed(2)} (${changePercent}% change)`;
        }
      }
    }

    // Fetch S&P 500 data
    const spyResponse = await fetch(
      'https://query1.finance.yahoo.com/v8/finance/chart/%5EGSPC?interval=1d&range=1d',
      { headers: { 'User-Agent': 'Mozilla/5.0' } }
    );
    
    let spData = "";
    if (spyResponse.ok) {
      const spJson = await spyResponse.json();
      const result = spJson?.chart?.result?.[0];
      if (result) {
        const price = result.meta?.regularMarketPrice;
        const prevClose = result.meta?.previousClose;
        if (price && prevClose) {
          const changePercent = ((price - prevClose) / prevClose * 100).toFixed(2);
          spData = `S&P 500: ${price.toFixed(2)} (${changePercent}% change)`;
        }
      }
    }

    if (ftseData || spData) {
      console.log("Market data fetched successfully:", ftseData, spData);
      return `CURRENT DATE: ${formattedDate}

LIVE MARKET DATA (use for realistic values):
- ${ftseData || "FTSE 100: data unavailable"}
- ${spData || "S&P 500: data unavailable"}
- Use these market trends to inform investment performance figures
- For UK £ amounts, use realistic variations based on current market conditions

REALISTIC VALUE GUIDELINES:
- UK Pension Values: £30,000 - £500,000 (typical range)
- Investment Returns: 3% - 8% annually (realistic long-term)
- Inflation Rates: 2% - 4% (typical assumption)
- Risk-Free Rates: 4% - 5% (current UK rates)`;
    }
  } catch (error) {
    console.log("Error fetching market data:", error);
  }

  // Fallback if market data unavailable
  return `CURRENT DATE: ${formattedDate}

REALISTIC VALUE GUIDELINES (use these ranges for professional financial documents):
- UK Pension Values: £30,000 - £500,000 (typical range)
- Investment Returns: 3% - 8% annually (realistic long-term)
- Inflation Rates: 2% - 4% (typical assumption)
- Risk-Free Rates: 4% - 5% (current UK rates)
- Annuity Rates: 5% - 7% (depending on age/type)
- Percentages: Use professional variations (e.g., 42.3% instead of round 40%)
- Monetary amounts: Use realistic figures with appropriate precision (e.g., £43,567 not £40,000)`;
}

async function generateAllReplacements(
  instructionPrompt: string,
  sections: HighlightedSection[],
  marketData: string
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

  const marketContext = marketData ? `\n\n${marketData}` : "";

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
            content: `You are a professional document editor replacing placeholder text in financial documents.

CRITICAL RULES:
1. You MUST generate NEW replacement text for EVERY section - NEVER return the original text unchanged
2. For author/preparer/reviewer names: ALWAYS use "Romit Acharya"
3. For numeric values (amounts, percentages): Generate realistic professional estimates appropriate for financial documents
4. For dates: Use current/recent dates as appropriate
5. Match the professional tone and style of the document context
6. Do NOT change client names, beneficiary names, or company names mentioned
7. Keep the replacement similar in length unless more detail improves clarity

EXAMPLES:
- "Roshan" → "Romit Acharya"
- "Prepared by [Name]" → "Prepared by Romit Acharya"
- "40%" → "42%" (generate a realistic variation)
- "£40,233" → "£43,500" (generate a realistic amount)
- "[Date]" → "15th January 2024"`
          },
        {
            role: "user",
            content: `DOCUMENT CONTEXT AND INSTRUCTIONS:\n${instructionPrompt}${marketContext}\n\nHIGHLIGHTED SECTIONS REQUIRING REPLACEMENT:\n${sectionsText}\n\nIMPORTANT: You MUST provide a NEW, DIFFERENT replacement for each section. Do not return the original text.`
          }
        ],
        tools: [
          {
            type: "function",
            function: {
              name: "provide_replacements",
              description: "Provide NEW replacement text for each highlighted section. Every replacement MUST be different from the original.",
              parameters: {
                type: "object",
                properties: {
                  replacements: {
                    type: "array",
                    description: "Array of replacements, one for each section",
                    items: {
                      type: "object",
                      properties: {
                        section_number: { type: "number", description: "The section number (1-indexed)" },
                        replacement_text: { type: "string", description: "The NEW replacement text - must be different from original" }
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
        temperature: 0.8
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
    // Escape special regex characters in the fullMatch for safe replacement
    const escapedMatch = escapeRegExp(replacement.fullMatch);
    
    // Create the new run with new text AND remove highlight formatting
    let newRun = replacement.fullMatch;
    
    // Replace the text content
    newRun = newRun.replace(
      /<w:t([^>]*)>[^<]*<\/w:t>/g,
      `<w:t$1>${escapeXml(replacement.newText)}</w:t>`
    );
    
    // Remove highlight formatting so the replaced text is no longer highlighted
    newRun = newRun.replace(/<w:highlight[^>]*\/>/g, '');
    newRun = newRun.replace(/<w:highlight[^>]*>.*?<\/w:highlight>/g, '');
    
    // Use regex for more reliable matching
    const regex = new RegExp(escapedMatch, 'g');
    modifiedXml = modifiedXml.replace(regex, newRun);
  }
  
  return modifiedXml;
}

function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

    console.log("=== EXTRACTED INSTRUCTION PROMPT ===");
    console.log(instructionPrompt);
    console.log("=== END INSTRUCTION PROMPT ===");

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

    // Fetch market data for realistic values
    console.log("Fetching current market data...");
    const marketData = await fetchMarketData();

    // Generate all replacements in a single batched API call
    console.log(`Processing ${highlightedSections.length} highlighted sections in a single batch request`);
    const replacementMap = await generateAllReplacements(instructionPrompt, highlightedSections, marketData);
    
    console.log(`Received ${replacementMap.size} replacements from AI`);
    
    const replacements: ProcessedReplacement[] = highlightedSections.map((section, index) => {
      const newText = replacementMap.get(index) || section.text;
      console.log(`Section ${index}: "${section.text.substring(0, 30)}..." -> "${newText.substring(0, 30)}..."`);
      return {
        originalText: section.text,
        newText: newText,
        fullMatch: section.fullMatch
      };
    });

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
