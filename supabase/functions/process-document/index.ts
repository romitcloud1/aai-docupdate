import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import JSZip from "https://esm.sh/jszip@3.10.1";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-supabase-client-platform, x-supabase-client-platform-version, x-supabase-client-runtime, x-supabase-client-runtime-version",
  "Access-Control-Expose-Headers": "Content-Disposition",
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

interface PieChartData {
  imagePath: string;
  growthPercent: number;
  defensivePercent: number;
  labels: string[];
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
2. For author/preparer/reviewer names in signature sections: ALWAYS use "Romit Acharya"
3. For numeric values (amounts, percentages): Generate realistic professional estimates appropriate for financial documents
4. For dates: Use CURRENT date from the market data provided (today's date)
5. Match the professional tone and style of the document context
6. ABSOLUTELY PRESERVE client names, beneficiary names, customer names - these are the people RECEIVING the document, not the authors
7. In "Client Declaration" or signature sections, the CLIENT NAME must stay as the original name - only replace the ADVISER/PREPARER name with "Romit Acharya"
8. Keep the replacement similar in length unless more detail improves clarity
9. For URLs/hyperlinks: Keep them clean without special XML characters - return plain URLs

EXAMPLES:
- "Prepared by [Name]" → "Prepared by Romit Acharya" (this is the preparer/adviser)
- "Client: [Name]" → Keep the original client name unchanged
- "Reviewed by Roshan" → "Reviewed by Romit Acharya"
- "40%" → "42%" (generate a realistic variation)
- "£40,233" → "£43,500" (generate a realistic amount)
- For dates: Use the CURRENT DATE from market data (not hardcoded old dates)`
          },
        {
            role: "user",
            content: `DOCUMENT CONTEXT AND INSTRUCTIONS:\n${instructionPrompt}${marketContext}\n\nHIGHLIGHTED SECTIONS REQUIRING REPLACEMENT:\n${sectionsText}\n\nIMPORTANT: 
- You MUST provide a NEW, DIFFERENT replacement for each section. Do not return the original text.
- PRESERVE all client/customer names - only change author/preparer/adviser names to "Romit Acharya"
- Use TODAY'S DATE for any date replacements (see market data for current date)
- Keep URLs clean and plain without XML encoding`
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
  // Check if text looks like a URL - if so, only escape the minimum required
  const isUrl = /^https?:\/\//.test(text.trim());
  
  if (isUrl) {
    // For URLs, only escape < and > which would break XML structure
    // Keep & as-is since URLs commonly contain them for query params
    return text
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }
  
  // For regular text, do full XML escaping
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Extract percentage data from replacements for pie chart generation
function extractPercentageDataFromReplacements(replacements: ProcessedReplacement[], originalDocXml: string): PieChartData | null {
  console.log(`Analyzing ${replacements.length} replacements for percentage data`);
  
  let growthPercent: number | null = null;
  let defensivePercent: number | null = null;

  // Collect all percentage values from replacements
  const allPercentages: { value: number; newText: string; originalText: string }[] = [];
  
  for (const replacement of replacements) {
    // Look for percentage patterns in the new text
    const percentMatches = replacement.newText.match(/(\d+\.?\d*)\s*%/g);
    if (percentMatches) {
      for (const match of percentMatches) {
        const val = parseFloat(match);
        if (val > 0 && val <= 100) {
          allPercentages.push({
            value: val,
            newText: replacement.newText,
            originalText: replacement.originalText
          });
        }
      }
    }
  }

  console.log(`Found ${allPercentages.length} percentage values in replacements`);
  
  // Log all found percentages for debugging
  for (const p of allPercentages) {
    console.log(`  Percentage: ${p.value}% from "${p.newText.substring(0, 30)}..."`);
  }

  // Strategy 1: Look for percentages near equity/growth or bond/defensive keywords in the original context
  for (const pct of allPercentages) {
    // Get context around the original text in the document
    const originalPos = originalDocXml.indexOf(pct.originalText);
    if (originalPos !== -1) {
      const start = Math.max(0, originalPos - 300);
      const end = Math.min(originalDocXml.length, originalPos + 300);
      const context = originalDocXml.substring(start, end).toLowerCase();
      
      // Check for growth/equity keywords
      if (/growth|equity|equities|stock|shares|risk/i.test(context)) {
        if (growthPercent === null) {
          growthPercent = pct.value;
          console.log(`Found growth/equity percentage: ${pct.value}%`);
        }
      }
      
      // Check for defensive/bond keywords  
      if (/defensive|bond|bonds|fixed|gilt|cash|income/i.test(context)) {
        if (defensivePercent === null) {
          defensivePercent = pct.value;
          console.log(`Found defensive/bond percentage: ${pct.value}%`);
        }
      }
    }
  }

  // Strategy 2: If we found one, calculate the other
  if (growthPercent !== null && defensivePercent === null) {
    defensivePercent = 100 - growthPercent;
    console.log(`Calculated defensive: ${defensivePercent}%`);
  } else if (defensivePercent !== null && growthPercent === null) {
    growthPercent = 100 - defensivePercent;
    console.log(`Calculated growth: ${growthPercent}%`);
  }

  // Strategy 3: Look for pairs that sum to 100
  if (growthPercent === null && defensivePercent === null) {
    // Look for two percentages that add up to ~100
    for (let i = 0; i < allPercentages.length; i++) {
      for (let j = i + 1; j < allPercentages.length; j++) {
        const sum = allPercentages[i].value + allPercentages[j].value;
        if (Math.abs(sum - 100) < 2) {
          // Found a likely pair
          const val1 = allPercentages[i].value;
          const val2 = allPercentages[j].value;
          
          // Assume larger is equities (common in balanced portfolios)
          growthPercent = Math.max(val1, val2);
          defensivePercent = Math.min(val1, val2);
          console.log(`Found pair summing to 100: ${growthPercent}% + ${defensivePercent}%`);
          break;
        }
      }
      if (growthPercent !== null) break;
    }
  }

  // Strategy 4: Look for specific patterns like "61% in equities" in the document
  const allocationPattern = /(\d+\.?\d*)\s*%?\s*(?:in\s+)?(?:equities|growth|stocks)/i;
  const bondPattern = /(\d+\.?\d*)\s*%?\s*(?:in\s+)?(?:bonds|defensive|fixed)/i;
  
  if (growthPercent === null) {
    const eqMatch = originalDocXml.match(allocationPattern);
    if (eqMatch) {
      growthPercent = parseFloat(eqMatch[1]);
      console.log(`Found equity allocation from document pattern: ${growthPercent}%`);
    }
  }
  
  if (defensivePercent === null) {
    const bondMatch = originalDocXml.match(bondPattern);
    if (bondMatch) {
      defensivePercent = parseFloat(bondMatch[1]);
      console.log(`Found bond allocation from document pattern: ${defensivePercent}%`);
    }
  }

  // Final calculation if we have one value
  if (growthPercent !== null && defensivePercent === null) {
    defensivePercent = 100 - growthPercent;
  } else if (defensivePercent !== null && growthPercent === null) {
    growthPercent = 100 - defensivePercent;
  }

  if (growthPercent !== null && defensivePercent !== null) {
    console.log(`Final pie chart data: Equities ${growthPercent}%, Bonds ${defensivePercent}%`);
    return {
      imagePath: "",
      growthPercent,
      defensivePercent,
      labels: ["Equities", "Bonds"]
    };
  }

  console.log("Could not determine percentage data for pie chart");
  return null;
}

// Generate a new pie chart using AI image generation
async function generatePieChart(data: PieChartData): Promise<Uint8Array | null> {
  const apiKey = Deno.env.get("LOVABLE_API_KEY");
  
  if (!apiKey) {
    console.log("LOVABLE_API_KEY not available for pie chart generation");
    return null;
  }

  const prompt = `Create a simple, clean 2D pie chart image with exactly 2 slices:
- First slice: ${data.growthPercent.toFixed(1)}% - use solid BLUE color (#4472C4)
- Second slice: ${data.defensivePercent.toFixed(1)}% - use solid ORANGE color (#ED7D31)

CRITICAL REQUIREMENTS:
- Simple 2D pie chart, NO 3D effects
- Clean white background
- Show percentage values directly ON each slice in white/black text for readability
- Small legend below showing: Blue = "Equities", Orange = "Bonds"
- Professional financial document style
- No fancy effects, shadows or gradients - keep it flat and simple
- Make sure the proportions are accurate to the percentages given
- Square image format, suitable for embedding in Word document`;

  try {
    console.log(`Generating pie chart with Equities: ${data.growthPercent}%, Bonds: ${data.defensivePercent}%`);
    
    const response = await fetch("https://ai.gateway.lovable.dev/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "google/gemini-3-pro-image-preview",
        messages: [
          {
            role: "user",
            content: prompt
          }
        ],
        modalities: ["image", "text"]
      })
    });

    if (!response.ok) {
      console.log(`Pie chart generation failed: ${response.status}`);
      const errText = await response.text();
      console.log(`Error details: ${errText}`);
      return null;
    }

    const result = await response.json();
    const imageData = result.choices?.[0]?.message?.images?.[0]?.image_url?.url;
    
    if (!imageData) {
      console.log("No image returned from AI");
      return null;
    }

    // Extract base64 data from data URL
    const base64Match = imageData.match(/^data:image\/\w+;base64,(.+)$/);
    if (!base64Match) {
      console.log("Invalid image data format");
      return null;
    }

    // Decode base64 to Uint8Array
    const binaryString = atob(base64Match[1]);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    
    console.log(`Successfully generated pie chart image (${bytes.length} bytes)`);
    return bytes;
    
  } catch (error) {
    console.log("Error generating pie chart:", error);
    return null;
  }
}

// Find pie chart images in the document and replace them
async function replacePieCharts(
  zip: JSZip,
  documentXml: string,
  modifiedXml: string,
  replacements: ProcessedReplacement[]
): Promise<void> {
  // Extract the new percentage data from the actual AI replacements
  // Pass the original document XML for context analysis
  const percentageData = extractPercentageDataFromReplacements(replacements, documentXml);
  
  if (!percentageData) {
    console.log("No percentage data found in replacements for pie chart generation");
    return;
  }

  console.log(`Using replacement percentages: Equities ${percentageData.growthPercent}%, Bonds ${percentageData.defensivePercent}%`);

  // Find all images in the document
  const mediaFolder = zip.folder("word/media");
  if (!mediaFolder) {
    console.log("No media folder found in document");
    return;
  }

  // Get list of image files
  const imageFiles: string[] = [];
  zip.forEach((relativePath, file) => {
    if (relativePath.startsWith("word/media/") && !file.dir) {
      const ext = relativePath.toLowerCase();
      if (ext.endsWith(".png") || ext.endsWith(".jpg") || ext.endsWith(".jpeg")) {
        imageFiles.push(relativePath);
      }
    }
  });

  console.log(`Found ${imageFiles.length} images in document`);

  if (imageFiles.length === 0) {
    return;
  }

  // Generate the new pie chart
  const newPieChartData = await generatePieChart(percentageData);
  
  if (!newPieChartData) {
    console.log("Failed to generate new pie chart");
    return;
  }

  // More selective pie chart detection - avoid replacing logos
  // Look for images that appear NEAR percentage/allocation text in the document
  
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (!relsFile) {
    console.log("No document.xml.rels found");
    return;
  }

  const relsContent = await relsFile.async("text");
  
  // Pattern to detect text near pie charts - look for allocation/investment language
  const pieChartContextPattern = /(?:asset\s*allocation|investment\s*mix|portfolio\s*breakdown|equities.*bonds|growth.*defensive|\d+\.?\d*%\s*(?:equities|bonds|growth|defensive))/gi;
  
  // Find the approximate position of pie chart context in the document
  const contextMatch = pieChartContextPattern.exec(modifiedXml);
  if (!contextMatch) {
    console.log("No pie chart context found in document - skipping image replacement to preserve logos");
    return;
  }
  
  const contextPosition = contextMatch.index;
  console.log(`Found pie chart context at position ${contextPosition}`);
  
  // Find the best candidate image for pie chart replacement
  // Prefer images that are close to allocation context but not logos
  let bestCandidate: { path: string; distance: number; position: number } | null = null;
  
  for (const imagePath of imageFiles) {
    const imageName = imagePath.split("/").pop() || "";
    
    // Skip images that are likely logos (often named logo, header, or appear very early in doc)
    if (imageName.toLowerCase().includes('logo') || 
        imageName.toLowerCase().includes('header') ||
        imageName.toLowerCase().includes('footer') ||
        imageName.toLowerCase().includes('signature')) {
      console.log(`Skipping likely logo/header image: ${imagePath}`);
      continue;
    }
    
    const imageIdMatch = relsContent.match(new RegExp(`Id="([^"]+)"[^>]*Target="media/${imageName}"`));
    
    if (imageIdMatch) {
      const rId = imageIdMatch[1];
      const imageUsagePattern = new RegExp(`<a:blip[^>]*r:embed="${rId}"`, 'gi');
      const imageMatch = imageUsagePattern.exec(modifiedXml);
      
      if (imageMatch) {
        const imagePosition = imageMatch.index;
        const distanceFromContext = Math.abs(imagePosition - contextPosition);
        
        console.log(`Image ${imageName} at position ${imagePosition}, distance from context: ${distanceFromContext}`);
        
        // Skip images at the very start (likely logo/header area - first 3000 chars)
        if (imagePosition < 3000) {
          console.log(`Skipping ${imageName} - appears to be in header area`);
          continue;
        }
        
        // Track the best candidate (closest to allocation context)
        if (bestCandidate === null || distanceFromContext < bestCandidate.distance) {
          bestCandidate = { path: imagePath, distance: distanceFromContext, position: imagePosition };
        }
      }
    }
  }
  
  // Replace the best candidate if found (within reasonable distance - 150000 chars)
  if (bestCandidate && bestCandidate.distance < 150000) {
    console.log(`Replacing pie chart image: ${bestCandidate.path} (distance: ${bestCandidate.distance})`);
    zip.file(bestCandidate.path, newPieChartData);
    console.log("Pie chart replaced successfully");
  } else if (bestCandidate) {
    console.log(`Best candidate ${bestCandidate.path} too far (${bestCandidate.distance} chars)`);
  } else {
    console.log("No suitable pie chart image found");
  }
}

serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const formData = await req.formData();
    const instructionFile = formData.get("instructionPrompt");
    const clientDataFile = formData.get("clientData");
    
    // Get original filename for the output
    let originalFileName = "document";
    if (clientDataFile && typeof clientDataFile === 'object' && 'name' in clientDataFile) {
      const fullName = (clientDataFile as File).name;
      // Remove .docx extension if present
      originalFileName = fullName.replace(/\.docx$/i, '');
    }

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
    
    // Attempt to regenerate pie charts with new data from actual replacements
    console.log("Checking for pie charts to regenerate...");
    await replacePieCharts(clientZip, documentXml, modifiedXml, replacements);
    
    const outputBuffer = await clientZip.generateAsync({ 
      type: "arraybuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 }
    });

    // Generate output filename with original name + current date (ddmmyyyy)
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const dateStr = `${day}${month}${year}`;
    const outputFileName = `${originalFileName}_${dateStr}.docx`;
    
    console.log(`Output filename: ${outputFileName}`);

    // Return the modified document
    return new Response(outputBuffer, {
      headers: {
        ...corsHeaders,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${outputFileName}"`
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
