export async function POST(request) {
try {
const { rawText, myModels = [] } = await request.json();
const hints = myModels.length
? '\n\nREFERENCE — my known model names (use these exact spellings when you can match them):\n' +
[...new Set(myModels.map(m => m.model))].slice(0, 80).join('\n')
: '';

const prompt = `You are a smartphone price list parser for a B2B mobile dealer in India.

Extract every smartphone model + variant + price from the WhatsApp message below.

STOCK TYPE RULES:


DEFAULT is FRESH. Add "(Fresh)" suffix to every item unless Activated is explicitly stated.
Only "(Activated)" if word Activated appears as section header or inline next to model.
NEVER leave suffix blank — always add (Fresh) or (Activated).


BRAND CONTEXT:


Brand-only lines set brand context for lines below. Combine brand + model.
"M7+" = "M7 Plus", "C85x" = "C85X"


PRICE FORMATS:


@price, space+number, Xk (13k=13000), "/-" is decoration
Trailing "-N" after price = qty, strip it
"Hold" or "NA" = skip item entirely


VARIANT: "4/128","6/128" etc. Fix typos: "6/28"→"6/128"
MODEL: Keep full brand+model. No 5G unless in message. No variant/price in name. Strip (X620) codes.
SKIP: Hold/NA items, brand-only headers, blank lines.

OUTPUT: ONLY a JSON array. No markdown, no explanation, no code fences.
[{"model":"...","variant":"...","price":12345}]
${hints}

Message:

${rawText}
---`;

// ── Gemini API (free tier) ────────────────────────────────────────────────
const apiKey = process.env.GEMINI_API_KEY;
const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

const res = await fetch(url, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,        // low temperature = more precise extraction
      maxOutputTokens: 4000,
    },
  }),
});

const data = await res.json();

// Gemini response structure: data.candidates[0].content.parts[0].text
if (data.error) throw new Error(data.error.message || 'Gemini API error');
const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
const clean = text.replace(/```json|```/g, '').trim();

let items;
try {
  items = JSON.parse(clean);
} catch {
  // Try to salvage partial array if response was cut off
  const m = clean.match(/(\[.*)/s);
  if (m) {
    const partial = m[1].replace(/,\s*$/, '').replace(/,\s*\{[^}]*$/, '') + ']';
    items = JSON.parse(partial);
  } else {
    throw new Error('Could not parse Gemini response as JSON');
  }
}

return Response.json({ items });

} catch (err) {
console.error('Parse API error:', err);
return Response.json({ error: err.message }, { status: 500 });
}
}
