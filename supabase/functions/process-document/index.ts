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

// Extract all text runs from document with their XML matches for potential replacement
function extractAllTextRuns(documentXml: string): { text: string; context: string; fullMatch: string }[] {
  const textRuns: { text: string; context: string; fullMatch: string }[] = [];
  
  const runPattern = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
  
  let contextBuffer: string[] = [];
  let runMatch: RegExpExecArray | null;
  
  while ((runMatch = runPattern.exec(documentXml)) !== null) {
    const runContent = runMatch[1];
    const fullMatch = runMatch[0];
    
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
      
      textRuns.push({
        text: runText,
        context: contextBuffer.slice(-5).join(" "),
        fullMatch: fullMatch
      });
    }
  }
  
  return textRuns;
}

// Legacy function for backward compatibility - still extract highlighted sections if they exist
function extractHighlightedSections(documentXml: string): HighlightedSection[] {
  const highlightedSections: HighlightedSection[] = [];
  
  const runPattern = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
  
  let contextBuffer: string[] = [];
  let runMatch: RegExpExecArray | null;
  
  while ((runMatch = runPattern.exec(documentXml)) !== null) {
    const runContent = runMatch[1];
    const fullMatch = runMatch[0];
    
    const hasHighlight = /<w:highlight\b[^\/]*\/>/.test(runContent) || /<w:highlight\b[^>]*>/.test(runContent);
    
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

// AI-driven placeholder identification - reads instruction prompt and client data to find what to replace
async function identifyAndReplaceWithAI(
  instructionPrompt: string,
  documentText: string,
  allTextRuns: { text: string; context: string; fullMatch: string }[],
  marketData: string
): Promise<ProcessedReplacement[]> {
  const apiKey = Deno.env.get("LOVABLE_API_KEY");
  
  if (!apiKey) {
    throw new Error("LOVABLE_API_KEY is not configured");
  }

  const maxRetries = 5;
  const baseDelay = 2000;

  // Build a text representation of available runs for the AI to identify
  const runsWithIndices = allTextRuns.map((run, i) => 
    `[${i}] "${run.text}"`
  ).join("\n");

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
            content: `You are a professional document editor. Your task is to analyze an instruction prompt and a client document to identify which text elements need to be replaced and what they should be replaced with.

CRITICAL: Follow the instruction prompt AND automatically detect financial data that needs updating.

Your job:
1. Read the instruction prompt carefully - it tells you what to look for and replace
2. Examine the available text runs from the document
3. Identify which runs match the patterns/placeholders described in the instruction prompt
4. ALSO automatically identify and replace ANY financial data including:
   - Monetary amounts (£, $, €, etc.) - generate realistic updated values
   - Percentages (%) - generate realistic professional variations
   - Dates (especially old dates) - update to current date or appropriate future dates
   - Financial metrics, returns, rates, yields
   - Investment values, pension amounts, fund values
5. Generate appropriate replacement text

IMPORTANT RULES:
- Follow the instruction prompt's rules precisely
- For author/preparer/reviewer names in PREPARATION or REVIEW contexts: Use "Romit Acharya"
- NEVER replace names that appear AFTER salutations like "Yours sincerely", "Yours truly", "Yours faithfully", "Kind regards", "Best regards", "Warm regards", "With thanks", "Respectfully" - these are SIGNATORY names and must be PRESERVED as-is
- PRESERVE client names, beneficiary names, customer names - these are the people RECEIVING the document
- PRESERVE signatory names in letter closings - keep the original name exactly as it appears
- Use TODAY'S DATE for any date replacements (see market data for current date)
- For numeric values: Generate realistic professional estimates with slight variations (e.g., £43,567 not £40,000)
- For percentages: Use realistic professional variations (e.g., 42.3% instead of round 40%)
- AUTOMATICALLY replace ALL outdated financial figures even if not explicitly mentioned in instruction prompt`
          },
          {
            role: "user",
            content: `=== INSTRUCTION PROMPT (FOLLOW THIS RELIGIOUSLY) ===
${instructionPrompt}
${marketContext}

=== AVAILABLE TEXT RUNS FROM DOCUMENT ===
${runsWithIndices}

Based on the instruction prompt above, identify which text runs need to be replaced and provide the new text for each. Only include runs that actually need replacement according to the instructions.`
          }
        ],
        tools: [
          {
            type: "function",
            function: {
              name: "provide_replacements",
              description: "Identify which text runs need replacement and provide new text for each",
              parameters: {
                type: "object",
                properties: {
                  replacements: {
                    type: "array",
                    description: "Array of replacements for identified text runs",
                    items: {
                      type: "object",
                      properties: {
                        run_index: { type: "number", description: "The index of the text run to replace (from the list provided)" },
                        replacement_text: { type: "string", description: "The new text to replace it with" },
                        reason: { type: "string", description: "Brief reason why this matches a pattern from instruction prompt" }
                      },
                      required: ["run_index", "replacement_text", "reason"]
                    }
                  }
                },
                required: ["replacements"]
              }
            }
          }
        ],
        tool_choice: { type: "function", function: { name: "provide_replacements" } },
        temperature: 0.3
      })
    });

    if (response.ok) {
      const data = await response.json();
      const toolCall = data.choices[0]?.message?.tool_calls?.[0];
      
      if (toolCall?.function?.arguments) {
        const parsed = JSON.parse(toolCall.function.arguments);
        const results: ProcessedReplacement[] = [];
        
        for (const r of parsed.replacements) {
          const runIndex = r.run_index;
          if (runIndex >= 0 && runIndex < allTextRuns.length) {
            const run = allTextRuns[runIndex];
            console.log(`AI identified replacement [${runIndex}]: "${run.text.substring(0, 30)}..." -> "${r.replacement_text.substring(0, 30)}..." (${r.reason})`);
            results.push({
              originalText: run.text,
              newText: r.replacement_text,
              fullMatch: run.fullMatch
            });
          }
        }
        
        return results;
      }
      throw new Error("Invalid AI response format");
    }

    if (response.status === 429 || response.status === 402) {
      if (attempt < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, attempt);
        console.log(`Rate limited/quota, retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw new Error("AI service temporarily unavailable. Please try again in a few minutes.");
    }

    const errorText = await response.text();
    throw new Error(`AI API error: ${response.status} - ${errorText}`);
  }

  throw new Error("Max retries exceeded");
}

async function generateBatchReplacements(
  instructionPrompt: string,
  sections: HighlightedSection[],
  startIndex: number,
  marketData: string
): Promise<Map<number, string>> {
  const apiKey = Deno.env.get("LOVABLE_API_KEY");
  
  if (!apiKey) {
    throw new Error("LOVABLE_API_KEY is not configured");
  }

  const maxRetries = 5;
  const baseDelay = 2000;

  // Build prompt with batch sections (using global indices)
  const sectionsText = sections.map((s, i) => 
    `[Section ${startIndex + i + 1}]\nContext: ${s.context}\nHighlighted text to replace: "${s.text}"`
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
          // Convert back to global 0-indexed
          const globalIndex = r.section_number - 1;
          resultMap.set(globalIndex, r.replacement_text);
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
      if (attempt < maxRetries - 1) {
        const delay = baseDelay * Math.pow(2, attempt);
        console.log(`Credits/quota issue, retrying in ${delay}ms (attempt ${attempt + 1}/${maxRetries})`);
        await new Promise(resolve => setTimeout(resolve, delay));
        continue;
      }
      throw new Error("AI service temporarily unavailable. Please try again in a few minutes.");
    }

    const errorText = await response.text();
    throw new Error(`AI API error: ${response.status} - ${errorText}`);
  }

  throw new Error("Max retries exceeded");
}

async function generateAllReplacements(
  instructionPrompt: string,
  sections: HighlightedSection[],
  marketData: string
): Promise<Map<number, string>> {
  const BATCH_SIZE = 50; // Process 50 sections at a time to avoid token limits
  const allReplacements = new Map<number, string>();
  
  const totalBatches = Math.ceil(sections.length / BATCH_SIZE);
  console.log(`Processing ${sections.length} sections in ${totalBatches} batches of up to ${BATCH_SIZE}`);
  
  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const startIdx = batchIndex * BATCH_SIZE;
    const endIdx = Math.min(startIdx + BATCH_SIZE, sections.length);
    const batchSections = sections.slice(startIdx, endIdx);
    
    console.log(`Processing batch ${batchIndex + 1}/${totalBatches} (sections ${startIdx + 1}-${endIdx})`);
    
    const batchReplacements = await generateBatchReplacements(
      instructionPrompt,
      batchSections,
      startIdx,
      marketData
    );
    
    // Merge batch results into main map
    for (const [index, text] of batchReplacements) {
      allReplacements.set(index, text);
    }
    
    // Add a small delay between batches to avoid rate limiting
    if (batchIndex < totalBatches - 1) {
      console.log("Waiting 1s before next batch...");
      await new Promise(resolve => setTimeout(resolve, 1000));
    }
  }
  
  console.log(`Generated ${allReplacements.size} replacements total`);
  return allReplacements;
}

function replaceHighlightedText(
  documentXml: string,
  replacements: ProcessedReplacement[]
): string {
  let modifiedXml = documentXml;
  
  // Track unique text snippets that were replaced (to find their paragraphs later)
  const replacedTextSnippets: string[] = [];
  
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
    
    // Remove all forms of highlight formatting so the replaced text is no longer highlighted
    newRun = removeHighlightFormatting(newRun);
    
    // Use regex for more reliable matching
    const regex = new RegExp(escapedMatch, 'g');
    modifiedXml = modifiedXml.replace(regex, newRun);
    
    // Store escaped new text to find its paragraph
    replacedTextSnippets.push(escapeXml(replacement.newText));
  }
  
  // Second pass: Clean up highlights from entire paragraphs containing replaced text
  // This catches bullet points and other runs in the same paragraph
  modifiedXml = cleanupParagraphHighlights(modifiedXml, replacedTextSnippets);
  
  return modifiedXml;
}

// Helper to remove all highlight formatting from a string
function removeHighlightFormatting(xml: string): string {
  let result = xml;
  
  // Remove <w:highlight> tags (self-closing and with content)
  result = result.replace(/<w:highlight[^>]*\/>/g, '');
  result = result.replace(/<w:highlight[^>]*>.*?<\/w:highlight>/g, '');
  
  // Remove <w:shd> (shading) tags that create highlight effects
  result = result.replace(/<w:shd[^>]*\/>/g, '');
  result = result.replace(/<w:shd[^>]*>.*?<\/w:shd>/g, '');
  
  // Remove background color attributes
  result = result.replace(/w:fill="[^"]*"/g, '');
  
  return result;
}

// Clean highlights from all runs in paragraphs that contain replaced text
function cleanupParagraphHighlights(xml: string, replacedTexts: string[]): string {
  let result = xml;
  
  // Find all paragraphs
  const paragraphPattern = /<w:p\b[^>]*>[\s\S]*?<\/w:p>/g;
  
  result = result.replace(paragraphPattern, (paragraph) => {
    // Check if this paragraph contains any of our replaced text
    const containsReplacedText = replacedTexts.some(text => 
      paragraph.includes(text) || paragraph.includes(text.substring(0, 20))
    );
    
    if (containsReplacedText) {
      // Remove highlights from ALL runs in this paragraph, not just the one we replaced
      return removeHighlightFormatting(paragraph);
    }
    
    return paragraph;
  });
  
  return result;
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

  // PRIORITY: Look for explicit allocation patterns in the NEW replacement text first
  // These patterns are the most reliable as they directly state the allocation
  for (const replacement of replacements) {
    const text = replacement.newText.toLowerCase();
    
    // Look for explicit equity/growth allocations like "54.2% in equities" or "54.2% equities"
    const equityMatch = text.match(/(\d+\.?\d*)\s*%?\s*(?:in\s+)?(?:equit(?:y|ies)|growth|stocks)/);
    if (equityMatch && growthPercent === null) {
      growthPercent = parseFloat(equityMatch[1]);
      console.log(`DIRECT MATCH - Found equity percentage in replacement: ${growthPercent}% from "${replacement.newText.substring(0, 50)}..."`);
    }
    
    // Look for explicit bond/defensive allocations
    const bondMatch = text.match(/(\d+\.?\d*)\s*%?\s*(?:in\s+)?(?:bonds?|defensive|fixed\s*income)/);
    if (bondMatch && defensivePercent === null) {
      defensivePercent = parseFloat(bondMatch[1]);
      console.log(`DIRECT MATCH - Found bond percentage in replacement: ${defensivePercent}% from "${replacement.newText.substring(0, 50)}..."`);
    }
  }

  // If we found both, we're done
  if (growthPercent !== null && defensivePercent !== null) {
    console.log(`Final pie chart data from direct matches: Equities ${growthPercent}%, Bonds ${defensivePercent}%`);
    return {
      imagePath: "",
      growthPercent,
      defensivePercent,
      labels: ["Equities", "Bonds"]
    };
  }

  // Collect all percentage values from replacements for fallback strategies
  const allPercentages: { value: number; newText: string; originalText: string }[] = [];
  
  for (const replacement of replacements) {
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

  // Strategy 2: Look for percentages that appear near equity/bond context in the replacement text
  if (growthPercent === null || defensivePercent === null) {
    for (const pct of allPercentages) {
      const text = pct.newText.toLowerCase();
      const pctStr = pct.value.toString();
      
      // Check if this percentage appears in a growth/equity context
      if (growthPercent === null && 
          (text.includes('growth') || text.includes('equit') || text.includes('stock')) &&
          text.includes(pctStr)) {
        growthPercent = pct.value;
        console.log(`Context match - Found growth percentage: ${pct.value}%`);
      }
      
      // Check if this percentage appears in a defensive/bond context  
      if (defensivePercent === null && 
          (text.includes('defensive') || text.includes('bond') || text.includes('fixed')) &&
          text.includes(pctStr)) {
        defensivePercent = pct.value;
        console.log(`Context match - Found defensive percentage: ${pct.value}%`);
      }
    }
  }

  // Strategy 3: If we have one, calculate the other
  if (growthPercent !== null && defensivePercent === null) {
    defensivePercent = Math.round((100 - growthPercent) * 10) / 10;
    console.log(`Calculated defensive: ${defensivePercent}%`);
  } else if (defensivePercent !== null && growthPercent === null) {
    growthPercent = Math.round((100 - defensivePercent) * 10) / 10;
    console.log(`Calculated growth: ${growthPercent}%`);
  }

  // Strategy 4: Look for pairs that sum to ~100
  if (growthPercent === null && defensivePercent === null && allPercentages.length >= 2) {
    for (let i = 0; i < allPercentages.length; i++) {
      for (let j = i + 1; j < allPercentages.length; j++) {
        const sum = allPercentages[i].value + allPercentages[j].value;
        if (Math.abs(sum - 100) < 2) {
          const val1 = allPercentages[i].value;
          const val2 = allPercentages[j].value;
          
          // Assume larger is equities (common in balanced portfolios)
          growthPercent = Math.max(val1, val2);
          defensivePercent = Math.min(val1, val2);
          console.log(`Found complementary pair: ${growthPercent}% + ${defensivePercent}%`);
          break;
        }
      }
      if (growthPercent !== null) break;
    }
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

// Add a new pie chart image to the document (instead of replacing existing images)
async function addPieChartToDocument(
  zip: JSZip,
  documentXml: string,
  modifiedXml: string,
  replacements: ProcessedReplacement[]
): Promise<string> {
  // Extract the new percentage data from the actual AI replacements
  const percentageData = extractPercentageDataFromReplacements(replacements, documentXml);
  
  if (!percentageData) {
    console.log("No percentage data found in replacements for pie chart generation");
    return modifiedXml;
  }

  console.log(`Using replacement percentages: Equities ${percentageData.growthPercent}%, Bonds ${percentageData.defensivePercent}%`);

  // Generate the new pie chart
  const newPieChartData = await generatePieChart(percentageData);
  
  if (!newPieChartData) {
    console.log("Failed to generate new pie chart");
    return modifiedXml;
  }

  // Find the highest image number in word/media/
  let highestImageNum = 0;
  zip.forEach((relativePath, file) => {
    if (relativePath.startsWith("word/media/image") && !file.dir) {
      const match = relativePath.match(/image(\d+)/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > highestImageNum) highestImageNum = num;
      }
    }
  });

  const newImageNum = highestImageNum + 1;
  const newImagePath = `word/media/image${newImageNum}.png`;
  
  // Add the new image to the zip
  zip.file(newImagePath, newPieChartData);
  console.log(`Added new pie chart as ${newImagePath}`);

  // Update the relationships file
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (relsFile) {
    let relsContent = await relsFile.async("text");
    
    // Find highest rId
    const rIdMatches = relsContent.matchAll(/Id="rId(\d+)"/g);
    let highestRId = 0;
    for (const match of rIdMatches) {
      const num = parseInt(match[1], 10);
      if (num > highestRId) highestRId = num;
    }
    const newRId = `rId${highestRId + 1}`;
    
    // Add new relationship before closing tag
    const newRel = `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${newImageNum}.png"/>`;
    relsContent = relsContent.replace('</Relationships>', `${newRel}</Relationships>`);
    zip.file("word/_rels/document.xml.rels", relsContent);
    console.log(`Added relationship ${newRId} for new pie chart`);

    // Find where to insert the pie chart in the document (near allocation context)
    const allocationPattern = /(?:asset\s*allocation|investment\s*mix|portfolio\s*breakdown|equities.*bonds|growth.*defensive|\d+\.?\d*%\s*(?:in\s+)?(?:equities|bonds))/gi;
    const contextMatch = allocationPattern.exec(modifiedXml);
    
    if (contextMatch) {
      // Find the end of the paragraph containing this match
      const insertPosition = modifiedXml.indexOf('</w:p>', contextMatch.index);
      
      if (insertPosition > -1) {
        // Create a drawing element for the pie chart - insert AFTER the closing </w:p>
        const endOfParagraph = insertPosition + 6; // Length of "</w:p>" is 6 characters
        
        const drawingXml = `
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:drawing>
<wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
<wp:extent cx="3600000" cy="3600000"/>
<wp:docPr id="${newImageNum + 100}" name="AI Generated Pie Chart"/>
<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
<pic:nvPicPr><pic:cNvPr id="${newImageNum}" name="piechart.png"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${newRId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="3600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic>
</a:graphicData>
</a:graphic>
</wp:inline>
</w:drawing></w:r>
</w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Asset Allocation: ${percentageData.growthPercent}% Equities, ${percentageData.defensivePercent}% Bonds</w:t></w:r></w:p>`;
        
        // Insert after the paragraph containing allocation text
        modifiedXml = modifiedXml.slice(0, endOfParagraph) + drawingXml + modifiedXml.slice(endOfParagraph);
        console.log("Inserted new pie chart into document");
      }
    }
  }

  return modifiedXml;
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
    const documentText = extractTextFromXml(documentXml);
    
    // First, check for highlighted sections in the client data document
    const highlightedSections = extractHighlightedSections(documentXml);
    
    // Fetch market data for realistic values
    console.log("Fetching current market data...");
    const marketData = await fetchMarketData();
    
    let replacements: ProcessedReplacement[];

    if (highlightedSections.length > 0) {
      // HIGHLIGHTED MODE: Only modify highlighted sections in client data
      console.log(`Found ${highlightedSections.length} highlighted sections in client data - using HIGHLIGHTED MODE`);
      console.log("Only highlighted sections will be modified, non-highlighted content will be preserved.");
      
      const replacementMap = await generateAllReplacements(instructionPrompt, highlightedSections, marketData);
      
      console.log(`Received ${replacementMap.size} replacements from AI`);
      
      replacements = highlightedSections.map((section, index) => {
        const newText = replacementMap.get(index) || section.text;
        console.log(`Section ${index}: "${section.text.substring(0, 30)}..." -> "${newText.substring(0, 30)}..."`);
        return {
          originalText: section.text,
          newText: newText,
          fullMatch: section.fullMatch
        };
      });
    } else {
      // NO HIGHLIGHTS MODE: AI automatically detects and replaces financial data
      console.log("No highlighted sections found in client data - using AUTO-DETECT MODE");
      console.log("AI will automatically detect and replace financial data (amounts, percentages, dates, rates).");
      
      // Extract all text runs from the document
      const allTextRuns = extractAllTextRuns(documentXml);
      console.log(`Extracted ${allTextRuns.length} text runs for AI analysis`);
      
      // Let AI identify what needs to be replaced - including automatic financial data detection
      replacements = await identifyAndReplaceWithAI(instructionPrompt, documentText, allTextRuns, marketData);
      
      console.log(`AI identified ${replacements.length} replacements`);
      
      if (replacements.length === 0) {
        return new Response(
          JSON.stringify({ error: "AI could not identify any text to replace. The document may not contain any financial data or patterns matching the instruction prompt." }),
          { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
        );
      }
    }

    // Create the modified document
    const modifiedXml = replaceHighlightedText(documentXml, replacements);
    clientZip.file("word/document.xml", modifiedXml);
    
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
