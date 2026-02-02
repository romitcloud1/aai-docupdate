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

interface LineChartData {
  title: string;
  dataPoints: { label: string; value: number }[];
  yAxisLabel: string;
  trend: 'up' | 'down' | 'stable';
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

// Extract performance/trend data from replacements for line chart generation
function extractLineChartDataFromReplacements(replacements: ProcessedReplacement[], modifiedXml: string): LineChartData | null {
  console.log(`Analyzing replacements for line chart data`);
  
  // Look for performance figures, returns, or time-series data in the replacements
  const dataPoints: { label: string; value: number }[] = [];
  let title = "Portfolio Performance";
  let yAxisLabel = "Value (£)";
  let trend: 'up' | 'down' | 'stable' = 'stable';
  
  // Extract monetary values with context for performance tracking
  for (const replacement of replacements) {
    const text = replacement.newText;
    
    // Look for monetary amounts with labels like "Start: £X" or "Current: £Y"
    const valuePatterns = [
      /(?:start(?:ing)?|initial|opening)\s*(?:value)?[:\s]*£?([\d,]+\.?\d*)/gi,
      /(?:current|final|closing|end)\s*(?:value)?[:\s]*£?([\d,]+\.?\d*)/gi,
      /(?:year\s*\d+|q[1-4]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[:\s]*£?([\d,]+\.?\d*)/gi
    ];
    
    // Extract start value
    const startMatch = text.match(/(?:start(?:ing)?|initial|opening)\s*(?:value)?[:\s]*£?([\d,]+\.?\d*)/i);
    if (startMatch) {
      const val = parseFloat(startMatch[1].replace(/,/g, ''));
      if (val > 0 && !dataPoints.find(p => p.label === 'Start')) {
        dataPoints.push({ label: 'Start', value: val });
      }
    }
    
    // Extract current/end value
    const endMatch = text.match(/(?:current|final|closing|end(?:ing)?)\s*(?:value)?[:\s]*£?([\d,]+\.?\d*)/i);
    if (endMatch) {
      const val = parseFloat(endMatch[1].replace(/,/g, ''));
      if (val > 0 && !dataPoints.find(p => p.label === 'Current')) {
        dataPoints.push({ label: 'Current', value: val });
      }
    }
    
    // Extract percentage returns
    const returnMatch = text.match(/(\+?-?\d+\.?\d*)\s*%\s*(?:return|growth|gain|performance)/i);
    if (returnMatch) {
      const returnPct = parseFloat(returnMatch[1]);
      trend = returnPct > 2 ? 'up' : returnPct < -2 ? 'down' : 'stable';
    }
  }
  
  // If we found start/current values, create intermediate points
  if (dataPoints.length >= 2) {
    const start = dataPoints.find(p => p.label === 'Start')?.value || 0;
    const end = dataPoints.find(p => p.label === 'Current')?.value || 0;
    
    // Generate realistic intermediate points
    const years = ['2020', '2021', '2022', '2023', '2024', '2025'];
    const totalGrowth = end - start;
    
    // Create a realistic growth curve
    const generatedPoints: { label: string; value: number }[] = [];
    for (let i = 0; i < years.length; i++) {
      const progress = i / (years.length - 1);
      // Add some variation to make it look realistic
      const variance = (Math.sin(i * 1.5) * 0.05) * totalGrowth;
      const value = Math.round(start + (totalGrowth * progress) + variance);
      generatedPoints.push({ label: years[i], value });
    }
    
    trend = end > start ? 'up' : end < start ? 'down' : 'stable';
    
    console.log(`Generated line chart data with ${generatedPoints.length} points, trend: ${trend}`);
    return {
      title,
      dataPoints: generatedPoints,
      yAxisLabel,
      trend
    };
  }
  
  // Fallback: look for any large monetary values to create a performance chart
  const allMonetaryValues: { value: number; context: string }[] = [];
  for (const replacement of replacements) {
    const matches = replacement.newText.matchAll(/£([\d,]+\.?\d*)/g);
    for (const match of matches) {
      const val = parseFloat(match[1].replace(/,/g, ''));
      if (val > 10000) { // Only consider significant values
        allMonetaryValues.push({ value: val, context: replacement.newText.substring(0, 50) });
      }
    }
  }
  
  if (allMonetaryValues.length >= 2) {
    // Sort by value to create a progression
    allMonetaryValues.sort((a, b) => a.value - b.value);
    
    const years = ['2020', '2021', '2022', '2023', '2024', '2025'];
    const minVal = allMonetaryValues[0].value;
    const maxVal = allMonetaryValues[allMonetaryValues.length - 1].value;
    
    const generatedPoints: { label: string; value: number }[] = [];
    for (let i = 0; i < years.length; i++) {
      const progress = i / (years.length - 1);
      const value = Math.round(minVal + (maxVal - minVal) * progress);
      generatedPoints.push({ label: years[i], value });
    }
    
    console.log(`Generated fallback line chart with ${generatedPoints.length} points`);
    return {
      title: "Portfolio Value Over Time",
      dataPoints: generatedPoints,
      yAxisLabel: "Value (£)",
      trend: maxVal > minVal ? 'up' : 'down'
    };
  }
  
  console.log("Could not extract sufficient data for line chart");
  return null;
}

// Generate a new line chart using AI image generation
async function generateLineChart(data: LineChartData): Promise<Uint8Array | null> {
  const apiKey = Deno.env.get("LOVABLE_API_KEY");
  
  if (!apiKey) {
    console.log("LOVABLE_API_KEY not available for line chart generation");
    return null;
  }

  // Format data points for the prompt
  const dataDescription = data.dataPoints
    .map(p => `${p.label}: £${p.value.toLocaleString()}`)
    .join(', ');

  const trendColor = data.trend === 'up' ? 'green (#2E7D32)' : data.trend === 'down' ? 'red (#C62828)' : 'blue (#1565C0)';
  
  const prompt = `Create a professional line chart for a financial document:

DATA POINTS: ${dataDescription}

REQUIREMENTS:
- Clean, professional line chart with smooth curve connecting points
- X-axis: Years (${data.dataPoints.map(p => p.label).join(', ')})
- Y-axis: ${data.yAxisLabel} with proper scale showing £ values
- Line color: ${trendColor} (solid line, 2-3px thickness)
- Add small circular markers at each data point
- Include subtle gridlines for readability
- White/light gray background
- Professional font styling for labels
- Title at top: "${data.title}"
- Clean, minimal style suitable for Word document embedding
- Landscape/wide format (16:9 or similar)
- NO 3D effects, keep it flat and clean`;

  try {
    console.log(`Generating line chart: ${data.title} with ${data.dataPoints.length} data points`);
    
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
      console.log(`Line chart generation failed: ${response.status}`);
      const errText = await response.text();
      console.log(`Error details: ${errText}`);
      return null;
    }

    const result = await response.json();
    const imageData = result.choices?.[0]?.message?.images?.[0]?.image_url?.url;
    
    if (!imageData) {
      console.log("No image returned from AI for line chart");
      return null;
    }

    // Extract base64 data from data URL
    const base64Match = imageData.match(/^data:image\/\w+;base64,(.+)$/);
    if (!base64Match) {
      console.log("Invalid line chart image data format");
      return null;
    }

    // Decode base64 to Uint8Array
    const binaryString = atob(base64Match[1]);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    
    console.log(`Successfully generated line chart image (${bytes.length} bytes)`);
    return bytes;
    
  } catch (error) {
    console.log("Error generating line chart:", error);
    return null;
  }
}

// Find and replace line chart images in the document
async function replaceLineCharts(
  zip: JSZip,
  documentXml: string,
  modifiedXml: string,
  replacements: ProcessedReplacement[]
): Promise<void> {
  const lineChartData = extractLineChartDataFromReplacements(replacements, modifiedXml);
  
  if (!lineChartData) {
    console.log("No line chart data found in replacements");
    return;
  }

  console.log(`Using line chart data: ${lineChartData.dataPoints.length} points, trend: ${lineChartData.trend}`);

  // Find all images in the document
  const imageFiles: string[] = [];
  zip.forEach((relativePath, file) => {
    if (relativePath.startsWith("word/media/") && !file.dir) {
      const ext = relativePath.toLowerCase();
      if (ext.endsWith(".png") || ext.endsWith(".jpg") || ext.endsWith(".jpeg")) {
        imageFiles.push(relativePath);
      }
    }
  });

  if (imageFiles.length === 0) {
    console.log("No images found for line chart replacement");
    return;
  }

  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (!relsFile) {
    console.log("No document.xml.rels found");
    return;
  }

  const relsContent = await relsFile.async("text");

  // Look for line chart context - performance, trend, growth over time
  const lineChartContextPattern = /(?:performance\s*(?:chart|graph|over\s*time)|growth\s*(?:chart|trend)|portfolio\s*(?:value|growth)\s*over|investment\s*(?:performance|growth))/gi;
  
  let contextMatch = lineChartContextPattern.exec(modifiedXml);
  let contextPosition: number | null = null;
  
  if (contextMatch) {
    contextPosition = contextMatch.index;
    console.log(`Found line chart context at position ${contextPosition}`);
  } else {
    // Fallback: look for time-series language
    const timeSeriesPattern = /(?:over\s*the\s*(?:past|last)\s*(?:\d+)?\s*years?|year[- ]on[- ]year|historical\s*(?:performance|growth))/gi;
    const timeMatch = timeSeriesPattern.exec(modifiedXml);
    if (timeMatch) {
      contextPosition = timeMatch.index;
      console.log(`Found time-series context at position ${contextPosition}`);
    }
  }

  if (contextPosition === null) {
    console.log("No line chart context found - skipping line chart replacement");
    return;
  }

  // Generate the new line chart
  const newLineChartData = await generateLineChart(lineChartData);
  
  if (!newLineChartData) {
    console.log("Failed to generate new line chart");
    return;
  }

  // Find candidate images for line chart (different from pie chart selection)
  const imageCandidates: { path: string; position: number; name: string }[] = [];
  
  for (const imagePath of imageFiles) {
    const imageName = imagePath.split("/").pop() || "";
    
    // Skip logos and headers
    if (imageName.toLowerCase().includes('logo') || 
        imageName.toLowerCase().includes('header') ||
        imageName.toLowerCase().includes('footer')) {
      continue;
    }
    
    const imageIdMatch = relsContent.match(new RegExp(`Id="([^"]+)"[^>]*Target="media/${imageName}"`));
    
    if (imageIdMatch) {
      const rId = imageIdMatch[1];
      const imageUsagePattern = new RegExp(`<a:blip[^>]*r:embed="${rId}"`, 'gi');
      const imageMatch = imageUsagePattern.exec(modifiedXml);
      
      if (imageMatch) {
        const imagePosition = imageMatch.index;
        
        // Skip header area
        if (imagePosition < 10000) {
          continue;
        }
        
        imageCandidates.push({ path: imagePath, position: imagePosition, name: imageName });
        console.log(`Line chart candidate: ${imageName} at position ${imagePosition}`);
      }
    }
  }

  if (imageCandidates.length === 0) {
    console.log("No suitable line chart candidates found");
    return;
  }

  // Sort by position
  imageCandidates.sort((a, b) => a.position - b.position);

  // Find the image closest to line chart context but BEFORE the pie chart context
  // (Line charts typically appear earlier in the document showing historical performance)
  let bestCandidate: { path: string; position: number; name: string } | null = null;
  
  // Look for the first image that's within reasonable distance of the context
  for (const candidate of imageCandidates) {
    const distance = Math.abs(candidate.position - contextPosition);
    
    // Pick an image close to the performance/line chart context
    if (distance < 80000) {
      // Make sure this isn't the same image as the pie chart
      // Pie charts tend to be later in the document near allocation sections
      const pieContextPattern = /(?:pie\s*chart|asset\s*allocation|allocation\s*breakdown)/gi;
      const pieMatch = pieContextPattern.exec(modifiedXml);
      
      if (pieMatch) {
        const pieContextPos = pieMatch.index;
        // If this image is closer to pie context, skip it
        if (Math.abs(candidate.position - pieContextPos) < Math.abs(candidate.position - contextPosition)) {
          console.log(`Skipping ${candidate.name} - closer to pie chart context`);
          continue;
        }
      }
      
      bestCandidate = candidate;
      console.log(`Selected ${candidate.name} as line chart (distance: ${distance})`);
      break;
    }
  }

  if (bestCandidate) {
    console.log(`Replacing line chart image: ${bestCandidate.path}`);
    zip.file(bestCandidate.path, newLineChartData);
    console.log("Line chart replaced successfully");
  } else {
    console.log("No suitable line chart image found within distance threshold");
  }
}


async function replacePieCharts(
  zip: JSZip,
  documentXml: string,
  modifiedXml: string,
  replacements: ProcessedReplacement[]
): Promise<{ modifiedXml: string }> {
  // Extract the new percentage data from the actual AI replacements
  const percentageData = extractPercentageDataFromReplacements(replacements, documentXml);
  
  if (!percentageData) {
    console.log("No percentage data found in replacements for pie chart generation");
    return { modifiedXml };
  }

  console.log(`Using replacement percentages: Equities ${percentageData.growthPercent}%, Bonds ${percentageData.defensivePercent}%`);

  // Find all images in the document
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
    console.log("No images found in document");
    return { modifiedXml };
  }

  // Get relationships file to map rId to image files
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (!relsFile) {
    console.log("No document.xml.rels found");
    return { modifiedXml };
  }

  const relsContent = await relsFile.async("text");

  // Find allocation context in the document for pie chart identification
  const allocationPatterns = [
    /(?:asset\s*allocation|equities.*bonds|growth.*defensive|portfolio\s*split|investment\s*mix)/gi,
    /(?:\d+\.?\d*\s*%\s*(?:in\s+)?(?:equit(?:y|ies)|bonds?|growth|defensive))/gi
  ];

  let allocationContextPosition: number | null = null;
  
  for (const pattern of allocationPatterns) {
    const match = pattern.exec(modifiedXml);
    if (match) {
      allocationContextPosition = match.index;
      console.log(`Found allocation context "${match[0]}" at position ${allocationContextPosition}`);
      break;
    }
  }

  // Find candidate pie chart images
  const imageCandidates: { path: string; position: number; name: string }[] = [];
  
  for (const imagePath of imageFiles) {
    const imageName = imagePath.split("/").pop() || "";
    
    // Skip logos and headers
    if (imageName.toLowerCase().includes('logo') || 
        imageName.toLowerCase().includes('header') ||
        imageName.toLowerCase().includes('footer')) {
      console.log(`Skipping ${imageName} - logo/header/footer`);
      continue;
    }
    
    // Find the relationship ID for this image
    const imageIdMatch = relsContent.match(new RegExp(`Id="([^"]+)"[^>]*Target="media/${imageName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`));
    
    if (imageIdMatch) {
      const rId = imageIdMatch[1];
      
      // Find where this image is used in the document
      const imageUsagePattern = new RegExp(`<a:blip[^>]*r:embed="${rId}"`, 'gi');
      const imageMatch = imageUsagePattern.exec(modifiedXml);
      
      if (imageMatch) {
        const imagePosition = imageMatch.index;
        
        // Skip images in the header area (first 10000 chars)
        if (imagePosition < 10000) {
          console.log(`Skipping ${imageName} - in header area (position ${imagePosition})`);
          continue;
        }
        
        imageCandidates.push({ path: imagePath, position: imagePosition, name: imageName });
        console.log(`Pie chart candidate: ${imageName} at position ${imagePosition}`);
      }
    }
  }

  if (imageCandidates.length === 0) {
    console.log("No suitable pie chart candidates found");
    return { modifiedXml };
  }

  // Sort by position
  imageCandidates.sort((a, b) => a.position - b.position);

  // Select the best pie chart candidate
  // Strategy: Find the image closest to allocation context
  let bestCandidate: { path: string; position: number; name: string } | null = null;
  
  if (allocationContextPosition !== null) {
    // Find image closest to allocation context
    let minDistance = Infinity;
    for (const candidate of imageCandidates) {
      const distance = Math.abs(candidate.position - allocationContextPosition);
      if (distance < minDistance) {
        minDistance = distance;
        bestCandidate = candidate;
      }
    }
    console.log(`Selected ${bestCandidate?.name} as pie chart - closest to allocation context (distance: ${minDistance})`);
  } else {
    // Fallback: Pick the LAST non-header/logo image (pie charts often appear later)
    bestCandidate = imageCandidates[imageCandidates.length - 1];
    console.log(`Selected ${bestCandidate?.name} as pie chart (last candidate - no allocation context found)`);
  }

  if (!bestCandidate) {
    console.log("No pie chart image selected");
    return { modifiedXml };
  }

  // Generate the new pie chart
  const newPieChartData = await generatePieChart(percentageData);
  
  if (!newPieChartData) {
    console.log("Failed to generate new pie chart");
    return { modifiedXml };
  }

  // REPLACE the existing pie chart image file
  console.log(`Replacing pie chart image: ${bestCandidate.path}`);
  zip.file(bestCandidate.path, newPieChartData);
  console.log(`Pie chart replaced successfully: ${bestCandidate.name} with Equities ${percentageData.growthPercent}%, Bonds ${percentageData.defensivePercent}%`);
  
  return { modifiedXml };
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
    let modifiedXml = replaceHighlightedText(documentXml, replacements);
    
    // Attempt to regenerate charts with new data from actual replacements
    console.log("Checking for line charts to regenerate...");
    await replaceLineCharts(clientZip, documentXml, modifiedXml, replacements);
    
    console.log("Checking for pie charts to regenerate (adding new chart)...");
    const pieChartResult = await replacePieCharts(clientZip, documentXml, modifiedXml, replacements);
    modifiedXml = pieChartResult.modifiedXml;
    
    // Save the final modified XML
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
