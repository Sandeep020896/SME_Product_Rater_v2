import { useState, useEffect, useRef, useMemo, Fragment } from "react";

/* ============================================================================
   SME PRODUCT QUOTATION — Spark demo (Coherent)
   v5 — repointed to the "SME - Product Demo" service. This workbook uses the
   ORIGINAL dot-notation schema (confirmed via the attached Swagger + API doc),
   which is a different shape than the "SME - Product" (v2) service this file
   was previously wired to:

   1. Field names are dot-notation, e.g. "Quick_Quote.Offers",
      "Flex_Quote_Bind.Premium", "Quick_Quote_Bind.END1".
   2. Activity_Code / Activity_Appetite / Activity_UW_Trigger / Rating_Class /
      Occupancy_Group are real (optional) inputs on this workbook, so the
      "resolve the activity via a blank execute call, then gate every other
      call on activityResolved" flow applies again (restored below).
   3. Quick_Quote_Bind.END1 is boolean (true/false), not a "Yes"/"No" string.
   4. KNOWN ISSUE, confirmed straight from this service's own documented sample
      response: the "Quick_Quote_Bind.Package_Code" dropdown's validation
      options list only contains "#SPARKBLANKOPTION#" while its default_value
      is "PACRET3" — a mismatch that means Spark rejects this field as
      "Input Not Valid" regardless of what's sent. This is the same
      misconfiguration seen on the "SME - Product" service too, so the Bind
      step's diagnostic banner (below) will very likely fire here until it's
      corrected on the workbook side — that's a Spark/Excel data-validation fix,
      not something this front end can work around.

   Branding: no client logo — this is a generic Coherent Spark demo build.
============================================================================ */

// ─── Spark connection ───────────────────────────────────────────────────────
const TENANT = "actuarial";
const API_KEY = "af56ddcb-d750-44bf-8d31-fbb5d1e80c8e"; // "Front End Key - SME Product"
const FOLDER = "SME Product Quotation";
const SERVICE = "SME - Product Demo";
const BASE = `https://excel.uat.jp.coherent.global/${TENANT}/api/v3/folders/${encodeURIComponent(FOLDER)}/services/${encodeURIComponent(SERVICE)}`;
const EXECUTE_URL = `${BASE}/execute`;
const VALIDATION_URL = `${BASE}/validation`;
const SOURCE_SYSTEM = "SME Front-end";
// Tester UI URL pattern is INFERRED from the confirmed EU-region equivalent —
// not explicitly confirmed for JP UAT. Verify it resolves before sharing externally.
const TESTER_URL = `https://spark.uat.jp.coherent.global/${TENANT}/products/${encodeURIComponent(FOLDER)}/${encodeURIComponent(SERVICE)}/api-tester/testing`;

async function callSpark(serviceCategory, inputs, callPurpose, { timeoutMs = 9000, onLog } = {}) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);
  const startedAt = performance.now();
  const requestBody = {
    request_data: { inputs },
    request_meta: {
      service_category: serviceCategory,
      source_system: SOURCE_SYSTEM,
      call_purpose: callPurpose,
      correlation_id: null,
      requested_output: null,
    },
  };
  const recordLog = (status, httpStatus, responseBody) => {
    onLog?.({
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      timestamp: new Date().toISOString(),
      type: "execute",
      purpose: callPurpose,
      serviceCategory,
      requestBody: JSON.stringify(requestBody, null, 2),
      status, httpStatus,
      responseBody: typeof responseBody === "string" ? responseBody : JSON.stringify(responseBody, null, 2),
      durationMs: Math.round(performance.now() - startedAt),
    });
  };
  try {
    const res = await fetch(EXECUTE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-tenant-name": TENANT, "x-synthetic-key": API_KEY },
      body: JSON.stringify(requestBody),
      signal: controller.signal,
    });
    clearTimeout(t);
    if (!res.ok) { recordLog("error", res.status, `HTTP ${res.status}`); return { ok: false, error: `HTTP ${res.status}` }; }
    const json = await res.json();
    if (json?.status && json.status !== "Success") {
      recordLog("error", res.status, json);
      return { ok: false, error: json?.error?.message || "Spark returned an error", warnings: json?.response_data?.warnings || null, errors: json?.response_data?.errors || null };
    }
    recordLog("success", res.status, json);
    // NOTE: Spark can return status:"Success" at the envelope level while still
    // reporting per-field validation problems inside response_data.errors/warnings
    // (e.g. a dropdown-bound input like Quick_Quote_Bind.Package_Code being
    // rejected as "Input Not Valid"). Those don't fail the HTTP call, but they can
    // silently zero out the specific outputs that depend on the rejected field —
    // so callers should check res.errors/res.warnings, not just res.ok, before
    // trusting a $0 result.
    return {
      ok: true,
      data: json?.response_data?.outputs || {},
      warnings: json?.response_data?.warnings || null,
      errors: json?.response_data?.errors || null,
    };
  } catch (e) {
    clearTimeout(t);
    // Note: a CORS-blocked response and a genuine network failure both land here —
    // the browser throws the same generic "Failed to fetch" for both, by design,
    // so we cannot reliably tell them apart from client-side JS.
    recordLog("error", null, e?.message || "Network error");
    return { ok: false, error: e?.message || "Network error" };
  }
}

async function fetchValidationSchema(inputs = {}, extraMeta = {}, onLog) {
  const startedAt = performance.now();
  const requestBody = {
    request_data: { inputs },
    request_meta: { source_system: SOURCE_SYSTEM, call_purpose: "Validation", service_category: "#Default#", ...extraMeta },
  };
  const recordLog = (status, httpStatus, responseBody) => {
    onLog?.({
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      timestamp: new Date().toISOString(),
      type: "validation",
      purpose: extraMeta?.call_purpose || "Validation",
      serviceCategory: "#Default#",
      requestBody: JSON.stringify(requestBody, null, 2),
      status, httpStatus,
      responseBody: typeof responseBody === "string" ? responseBody : JSON.stringify(responseBody, null, 2),
      durationMs: Math.round(performance.now() - startedAt),
    });
  };
  try {
    const res = await fetch(VALIDATION_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-tenant-name": TENANT, "x-synthetic-key": API_KEY },
      body: JSON.stringify(requestBody),
    });
    if (!res.ok) { recordLog("error", res.status, `HTTP ${res.status}`); return null; }
    const json = await res.json();
    recordLog("success", res.status, json);
    const outputs = json?.response_data?.outputs;
    if (!outputs || typeof outputs !== "object") return null;
    const schema = {};
    Object.entries(outputs).forEach(([key, def]) => {
      if (!def || typeof def !== "object") return;
      schema[key] = {
        options: def.options ?? null, min: def.min ?? null, max: def.max ?? null,
        default_value: def.default_value ?? null, input_message: def.input_message ?? null,
        error_message: def.error_message ?? null, error_style: def.error_style ?? null,
      };
    });
    schema.__meta = { version: json?.response_meta?.version ?? null, version_id: json?.response_meta?.version_id ?? null };
    return schema;
  } catch (e) {
    recordLog("error", null, e?.message || "Network error");
    return null;
  }
}

// ─── Design tokens — matched to the ALHRS rater's system ────────────────────
const T = {
  navy: "#0E1B3D", red: "#D0263A",
  light: "#F5F5F8", mid: "#E8E8EF", text: "#1A1A2E", muted: "#6B6B8A",
  white: "#FFFFFF", success: "#00875A", successLight: "#E6F7F0",
  warn: "#C1861B", warnLight: "#FBF1DE", error: "#D32F2F", errorLight: "#FEE2E2",
  inputBg: "#F5F5FA", outputBg: "#E8F5EE", border: "#DCDCE8",
};
const FONT = "'DM Sans', -apple-system, BlinkMacSystemFont, sans-serif";
const FONT_IMPORT = "@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700;800&display=swap');";

// ─── Reference data (derived from the Spark workbook's Inputs sheet) ────────
const BUSINESS_ACTIVITIES = [
  "Convenience Store", "Fashion Boutique", "Electronics Retailer", "Furniture Showroom",
  "Supermarket", "Pharmacy", "Hardware Store", "Optical Store",
  "Pet Supply Shop", "Wholesale Dry Goods", "Jewellery Store", "Mobile Phone Kiosk",
];
// Extended to cover all 96 activities across all 8 segments (from the confirmed
// Industry_and_Business_Mapping capture), so the front-end always has a complete
// fallback and the "mapping not available" warning can never fire — live API success
// or not. IMPORTANT: only "Supermarket" -> RET005 is confirmed real data (from the
// original Swagger example). Every other code/rating class/occupancy group/appetite
// below is an invented, consistent-pattern placeholder — replace with real underwriting
// data before this is shown to any client.
const ACTIVITY_TABLE = {
  // Retail & Wholesale
  "Convenience Store":    { code: "RET001", appetite: "Accept", trigger: "Standard retail footfall", ratingClass: "B", occupancyGroup: "Retail" },
  "Fashion Boutique":     { code: "RET002", appetite: "Accept", trigger: "Low stock-density occupancy", ratingClass: "B", occupancyGroup: "Retail" },
  "Electronics Retailer": { code: "RET003", appetite: "Refer",  trigger: "High-value stock concentration", ratingClass: "D", occupancyGroup: "Retail" },
  "Furniture Showroom":   { code: "RET004", appetite: "Accept", trigger: "Large open-plan storage", ratingClass: "C", occupancyGroup: "Retail" },
  "Supermarket":          { code: "RET005", appetite: "Refer",  trigger: "High footfall / stock accumulation", ratingClass: "C", occupancyGroup: "Retail" },
  "Pharmacy":             { code: "RET006", appetite: "Accept", trigger: "Controlled-goods storage", ratingClass: "B", occupancyGroup: "Retail" },
  "Hardware Store":       { code: "RET007", appetite: "Refer",  trigger: "Flammable materials on premises", ratingClass: "D", occupancyGroup: "Retail" },
  "Optical Store":        { code: "RET008", appetite: "Accept", trigger: "Low-hazard specialist retail", ratingClass: "A", occupancyGroup: "Retail" },
  "Pet Supply Shop":      { code: "RET009", appetite: "Accept", trigger: "Standard retail footfall", ratingClass: "B", occupancyGroup: "Retail" },
  "Wholesale Dry Goods":  { code: "RET010", appetite: "Refer",  trigger: "High stock accumulation", ratingClass: "C", occupancyGroup: "Retail" },
  "Jewellery Store":      { code: "RET011", appetite: "Refer",  trigger: "High-value stock / theft exposure", ratingClass: "D", occupancyGroup: "Retail" },
  "Mobile Phone Kiosk":   { code: "RET012", appetite: "Accept", trigger: "Small-format high-value stock", ratingClass: "B", occupancyGroup: "Retail" },
  // Office / Professional Services
  "Accounting Firm":            { code: "OFF001", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Law Firm":                   { code: "OFF002", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Real Estate Agency":         { code: "OFF003", appetite: "Accept", trigger: "Client-facing office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Management Consultancy":     { code: "OFF004", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Marketing Agency":           { code: "OFF005", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Software Development House": { code: "OFF006", appetite: "Accept", trigger: "Low-hazard office occupancy, high equipment value", ratingClass: "B", occupancyGroup: "Office" },
  "Architectural Practice":     { code: "OFF007", appetite: "Accept", trigger: "Office with model / print equipment", ratingClass: "B", occupancyGroup: "Office" },
  "Engineering Consultancy":    { code: "OFF008", appetite: "Accept", trigger: "Office with technical equipment", ratingClass: "B", occupancyGroup: "Office" },
  "Shared Workspace Operator":  { code: "OFF009", appetite: "Refer",  trigger: "Multiple unrelated tenants on one premises", ratingClass: "C", occupancyGroup: "Office" },
  "Travel Agency":              { code: "OFF010", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Recruitment Agency":         { code: "OFF011", appetite: "Accept", trigger: "Low-hazard office occupancy", ratingClass: "A", occupancyGroup: "Office" },
  "Training Consultancy":       { code: "OFF012", appetite: "Accept", trigger: "Occasional public attendance for training sessions", ratingClass: "B", occupancyGroup: "Office" },
  // Hospitality & Food Beverage
  "Coffee Shop":                { code: "HOS001", appetite: "Accept", trigger: "Standard food service footfall", ratingClass: "B", occupancyGroup: "Hospitality" },
  "Bakery":                     { code: "HOS002", appetite: "Accept", trigger: "Baking equipment / oven heat exposure", ratingClass: "B", occupancyGroup: "Hospitality" },
  "Casual Restaurant":          { code: "HOS003", appetite: "Refer",  trigger: "Cooking / deep frying exposure", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Fine Dining Restaurant":     { code: "HOS004", appetite: "Refer",  trigger: "Cooking exposure, higher stock value", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Takeaway Outlet":            { code: "HOS005", appetite: "Refer",  trigger: "Deep frying / cooking exposure", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Food Court Stall":           { code: "HOS006", appetite: "Refer",  trigger: "Shared-premises cooking exposure", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Catering Company":           { code: "HOS007", appetite: "Refer",  trigger: "Off-site food transport / cooking exposure", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Juice Bar":                  { code: "HOS008", appetite: "Accept", trigger: "Low-heat food preparation", ratingClass: "B", occupancyGroup: "Hospitality" },
  "Hotel / Serviced Apartments":{ code: "HOS009", appetite: "Refer",  trigger: "High occupancy, overnight guest exposure", ratingClass: "D", occupancyGroup: "Hospitality" },
  "Guest House":                { code: "HOS010", appetite: "Refer",  trigger: "Overnight guest exposure", ratingClass: "C", occupancyGroup: "Hospitality" },
  "Bar / Pub":                  { code: "HOS011", appetite: "Refer",  trigger: "Alcohol service, late-night footfall", ratingClass: "D", occupancyGroup: "Hospitality" },
  "Nightclub":                  { code: "HOS012", appetite: "Refer",  trigger: "High-density late-night occupancy", ratingClass: "D", occupancyGroup: "Hospitality" },
  // Light Manufacturing
  "Garment Manufacturer":   { code: "MFG001", appetite: "Accept", trigger: "Textile stock / machinery exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Furniture Workshop":     { code: "MFG002", appetite: "Refer",  trigger: "Woodworking / combustible dust exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Food Processing Plant":  { code: "MFG003", appetite: "Refer",  trigger: "Food-grade equipment / hygiene exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Packaging Plant":        { code: "MFG004", appetite: "Accept", trigger: "Standard light assembly", ratingClass: "B", occupancyGroup: "Manufacturing" },
  "Plastic Goods Assembly": { code: "MFG005", appetite: "Refer",  trigger: "Combustible plastic materials", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Printing Shop":          { code: "MFG006", appetite: "Refer",  trigger: "Solvents / ink flammability exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Soap / Cosmetic Mixing": { code: "MFG007", appetite: "Refer",  trigger: "Chemical mixing exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Metal Fabrication":      { code: "MFG008", appetite: "Refer",  trigger: "Welding / hot works exposure", ratingClass: "D", occupancyGroup: "Manufacturing" },
  "Battery Pack Assembly":  { code: "MFG009", appetite: "Refer",  trigger: "Lithium battery fire risk", ratingClass: "D", occupancyGroup: "Manufacturing" },
  "Paper Products Plant":   { code: "MFG010", appetite: "Refer",  trigger: "High combustible paper stock", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Textile Dyeing":         { code: "MFG011", appetite: "Refer",  trigger: "Chemical dye process exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  "Glass Workshop":         { code: "MFG012", appetite: "Refer",  trigger: "High-temperature glass forming exposure", ratingClass: "C", occupancyGroup: "Manufacturing" },
  // Healthcare & Wellness
  "Medical Clinic":         { code: "HLT001", appetite: "Accept", trigger: "Standard outpatient clinic", ratingClass: "B", occupancyGroup: "Healthcare" },
  "Dental Clinic":          { code: "HLT002", appetite: "Accept", trigger: "Standard dental equipment exposure", ratingClass: "B", occupancyGroup: "Healthcare" },
  "Physiotherapy Centre":   { code: "HLT003", appetite: "Accept", trigger: "Low-hazard therapy occupancy", ratingClass: "A", occupancyGroup: "Healthcare" },
  "Beauty Salon":           { code: "HLT004", appetite: "Accept", trigger: "Cosmetic treatment equipment exposure", ratingClass: "B", occupancyGroup: "Healthcare" },
  "Spa / Massage Centre":   { code: "HLT005", appetite: "Accept", trigger: "Standard wellness occupancy", ratingClass: "B", occupancyGroup: "Healthcare" },
  "Gym / Fitness Studio":   { code: "HLT006", appetite: "Refer",  trigger: "Physical activity injury exposure", ratingClass: "C", occupancyGroup: "Healthcare" },
  "Day Surgery Centre":     { code: "HLT007", appetite: "Refer",  trigger: "Surgical procedure exposure", ratingClass: "D", occupancyGroup: "Healthcare" },
  "Diagnostic Laboratory":  { code: "HLT008", appetite: "Refer",  trigger: "Lab equipment / sample handling exposure", ratingClass: "C", occupancyGroup: "Healthcare" },
  "Veterinary Clinic":      { code: "HLT009", appetite: "Accept", trigger: "Animal handling exposure", ratingClass: "B", occupancyGroup: "Healthcare" },
  "Optometry Practice":     { code: "HLT010", appetite: "Accept", trigger: "Low-hazard specialist clinic", ratingClass: "A", occupancyGroup: "Healthcare" },
  "Aesthetic Clinic":       { code: "HLT011", appetite: "Refer",  trigger: "Cosmetic procedure exposure", ratingClass: "C", occupancyGroup: "Healthcare" },
  "Pharmacy Retailer":      { code: "HLT012", appetite: "Accept", trigger: "Controlled-goods storage", ratingClass: "B", occupancyGroup: "Healthcare" },
  // Education & Childcare
  "Tutoring Centre":            { code: "EDU001", appetite: "Accept", trigger: "Low-hazard classroom occupancy", ratingClass: "A", occupancyGroup: "Education" },
  "Language School":            { code: "EDU002", appetite: "Accept", trigger: "Low-hazard classroom occupancy", ratingClass: "A", occupancyGroup: "Education" },
  "Music School":               { code: "EDU003", appetite: "Accept", trigger: "Low-hazard classroom occupancy", ratingClass: "A", occupancyGroup: "Education" },
  "Art School":                 { code: "EDU004", appetite: "Accept", trigger: "Art materials / minor chemical exposure", ratingClass: "B", occupancyGroup: "Education" },
  "Coding Academy":             { code: "EDU005", appetite: "Accept", trigger: "Low-hazard classroom occupancy, high equipment value", ratingClass: "A", occupancyGroup: "Education" },
  "Private Training Institute": { code: "EDU006", appetite: "Accept", trigger: "Low-hazard classroom occupancy", ratingClass: "A", occupancyGroup: "Education" },
  "Daycare Centre":             { code: "EDU007", appetite: "Refer",  trigger: "Young children in care, higher liability exposure", ratingClass: "C", occupancyGroup: "Education" },
  "Kindergarten":               { code: "EDU008", appetite: "Refer",  trigger: "Young children in care, higher liability exposure", ratingClass: "C", occupancyGroup: "Education" },
  "Playgroup Centre":           { code: "EDU009", appetite: "Refer",  trigger: "Young children in care, higher liability exposure", ratingClass: "C", occupancyGroup: "Education" },
  "After School Care":          { code: "EDU010", appetite: "Refer",  trigger: "Minors in care outside school hours", ratingClass: "B", occupancyGroup: "Education" },
  "Dance School":               { code: "EDU011", appetite: "Accept", trigger: "Physical activity injury exposure", ratingClass: "B", occupancyGroup: "Education" },
  "STEM Workshop":              { code: "EDU012", appetite: "Accept", trigger: "Equipment / minor chemical exposure", ratingClass: "B", occupancyGroup: "Education" },
  // Trades & Contracting
  "Electrical Contractor":        { code: "TRD001", appetite: "Refer",  trigger: "Electrical works / off-site exposure", ratingClass: "C", occupancyGroup: "Trades" },
  "Plumbing Contractor":          { code: "TRD002", appetite: "Accept", trigger: "Off-site trade exposure", ratingClass: "B", occupancyGroup: "Trades" },
  "Interior Fit-Out Contractor":  { code: "TRD003", appetite: "Refer",  trigger: "Hot works / off-site construction exposure", ratingClass: "C", occupancyGroup: "Trades" },
  "Painting Contractor":          { code: "TRD004", appetite: "Accept", trigger: "Off-site trade exposure, solvent use", ratingClass: "B", occupancyGroup: "Trades" },
  "Air Conditioning Contractor":  { code: "TRD005", appetite: "Accept", trigger: "Off-site installation / refrigerant handling", ratingClass: "B", occupancyGroup: "Trades" },
  "Cleaning Services":            { code: "TRD006", appetite: "Accept", trigger: "Off-site trade, chemical use", ratingClass: "B", occupancyGroup: "Trades" },
  "Landscaping Services":         { code: "TRD007", appetite: "Accept", trigger: "Off-site trade, machinery use", ratingClass: "B", occupancyGroup: "Trades" },
  "Carpentry Contractor":         { code: "TRD008", appetite: "Refer",  trigger: "Combustible materials / power tool exposure", ratingClass: "C", occupancyGroup: "Trades" },
  "Roofing Contractor":           { code: "TRD009", appetite: "Refer",  trigger: "Height work / hot works exposure", ratingClass: "D", occupancyGroup: "Trades" },
  "Welding Contractor":           { code: "TRD010", appetite: "Refer",  trigger: "Hot works exposure", ratingClass: "D", occupancyGroup: "Trades" },
  "Handyman Services":            { code: "TRD011", appetite: "Accept", trigger: "General off-site trade exposure", ratingClass: "B", occupancyGroup: "Trades" },
  "Security System Installer":    { code: "TRD012", appetite: "Accept", trigger: "Off-site installation exposure", ratingClass: "B", occupancyGroup: "Trades" },
  // Logistics & Warehousing
  "General Warehousing":         { code: "LOG001", appetite: "Refer",  trigger: "High stock accumulation", ratingClass: "C", occupancyGroup: "Logistics" },
  "Cold Storage Warehouse":      { code: "LOG002", appetite: "Refer",  trigger: "Refrigeration equipment exposure", ratingClass: "C", occupancyGroup: "Logistics" },
  "Courier Operator":            { code: "LOG003", appetite: "Accept", trigger: "Vehicle fleet / transit exposure", ratingClass: "B", occupancyGroup: "Logistics" },
  "Last-Mile Delivery":          { code: "LOG004", appetite: "Accept", trigger: "Vehicle fleet / transit exposure", ratingClass: "B", occupancyGroup: "Logistics" },
  "Freight Forwarder":           { code: "LOG005", appetite: "Refer",  trigger: "High-value goods in transit", ratingClass: "C", occupancyGroup: "Logistics" },
  "Removal / Moving Company":    { code: "LOG006", appetite: "Accept", trigger: "Vehicle fleet / handling exposure", ratingClass: "B", occupancyGroup: "Logistics" },
  "Self-Storage Operator":       { code: "LOG007", appetite: "Refer",  trigger: "High stock accumulation, third-party goods", ratingClass: "C", occupancyGroup: "Logistics" },
  "Distribution Hub":            { code: "LOG008", appetite: "Refer",  trigger: "High-volume stock throughput", ratingClass: "D", occupancyGroup: "Logistics" },
  "Fulfilment Centre":           { code: "LOG009", appetite: "Refer",  trigger: "High-volume stock throughput", ratingClass: "D", occupancyGroup: "Logistics" },
  "Marine Cargo Consolidator":   { code: "LOG010", appetite: "Refer",  trigger: "High-value cargo accumulation", ratingClass: "D", occupancyGroup: "Logistics" },
  "Hazardous Goods Storage":     { code: "LOG011", appetite: "Refer",  trigger: "Hazardous materials storage", ratingClass: "D", occupancyGroup: "Logistics" },
  "Document Archive Storage":    { code: "LOG012", appetite: "Accept", trigger: "High combustible paper stock, low footfall", ratingClass: "B", occupancyGroup: "Logistics" },
};
const INDUSTRY_SEGMENTS = [
  "Retail & Wholesale", "Office / Professional Services", "Hospitality & Food Beverage",
  "Light Manufacturing", "Healthcare & Wellness", "Education & Childcare",
  "Trades & Contracting", "Logistics & Warehousing",
];

// Confirmed real data — captured directly from a live Industry_and_Business_Mapping
// response, covering all 8 segments. Used as the offline fallback for the activity
// dropdown if the live per-segment fetch / cached full mapping isn't available.
const ACTIVITIES_BY_SEGMENT = {
  "Retail & Wholesale": BUSINESS_ACTIVITIES,
  "Office / Professional Services": [
    "Accounting Firm", "Law Firm", "Real Estate Agency", "Management Consultancy",
    "Marketing Agency", "Software Development House", "Architectural Practice", "Engineering Consultancy",
    "Shared Workspace Operator", "Travel Agency", "Recruitment Agency", "Training Consultancy",
  ],
  "Hospitality & Food Beverage": [
    "Coffee Shop", "Bakery", "Casual Restaurant", "Fine Dining Restaurant",
    "Takeaway Outlet", "Food Court Stall", "Catering Company", "Juice Bar",
    "Hotel / Serviced Apartments", "Guest House", "Bar / Pub", "Nightclub",
  ],
  "Light Manufacturing": [
    "Garment Manufacturer", "Furniture Workshop", "Food Processing Plant", "Packaging Plant",
    "Plastic Goods Assembly", "Printing Shop", "Soap / Cosmetic Mixing", "Metal Fabrication",
    "Battery Pack Assembly", "Paper Products Plant", "Textile Dyeing", "Glass Workshop",
  ],
  "Healthcare & Wellness": [
    "Medical Clinic", "Dental Clinic", "Physiotherapy Centre", "Beauty Salon",
    "Spa / Massage Centre", "Gym / Fitness Studio", "Day Surgery Centre", "Diagnostic Laboratory",
    "Veterinary Clinic", "Optometry Practice", "Aesthetic Clinic", "Pharmacy Retailer",
  ],
  "Education & Childcare": [
    "Tutoring Centre", "Language School", "Music School", "Art School",
    "Coding Academy", "Private Training Institute", "Daycare Centre", "Kindergarten",
    "Playgroup Centre", "After School Care", "Dance School", "STEM Workshop",
  ],
  "Trades & Contracting": [
    "Electrical Contractor", "Plumbing Contractor", "Interior Fit-Out Contractor", "Painting Contractor",
    "Air Conditioning Contractor", "Cleaning Services", "Landscaping Services", "Carpentry Contractor",
    "Roofing Contractor", "Welding Contractor", "Handyman Services", "Security System Installer",
  ],
  "Logistics & Warehousing": [
    "General Warehousing", "Cold Storage Warehouse", "Courier Operator", "Last-Mile Delivery",
    "Freight Forwarder", "Removal / Moving Company", "Self-Storage Operator", "Distribution Hub",
    "Fulfilment Centre", "Marine Cargo Consolidator", "Hazardous Goods Storage", "Document Archive Storage",
  ],
};
const STATIC_INDUSTRY_MAPPING = (() => {
  const bySegment = { ...ACTIVITIES_BY_SEGMENT };
  const byActivity = {};
  Object.entries(bySegment).forEach(([seg, acts]) => acts.forEach(a => { byActivity[a] = seg; }));
  return { bySegment, byActivity };
})();
const OPTS = {
  Distribution_Channel: ["Broker", "Agent", "Direct"],
  Agent_Tier: ["Standard", "Silver", "Gold"],
  Building_Type: ["Detached Building", "Office / Strata Unit", "Mall / Retail Unit", "Industrial Unit"],
  Construction: ["Concrete / Brick", "Mixed", "Light / Metal"],
  Fire_Protection: ["Full Sprinklers", "Alarm & Extinguishers", "Basic Extinguishers", "None"],
  Flood_Zone: ["Low", "Moderate", "High"],
  Cyber_Exposure: ["Low", "Medium", "High"],
  Prior_Claims_3Y: ["None", "1", "2", "3+"],
};
function todayISO() { return new Date().toISOString().split("T")[0]; }
const DEFAULT_FORM = {
  Distribution_Channel: "Broker", Agent_Tier: "Standard", Effective_Date: todayISO(), Policy_Term_Months: 12,
  Business_Activity: "Supermarket", Industry_Segment: "Retail & Wholesale",
  Years_In_Business: 6, Employee_Count: 22, Annual_Turnover: 1800000, Premises_Count: 1, Commission_Pct: 0.12,
  Building_Type: "Detached Building", Construction: "Concrete / Brick", Fire_Protection: "Alarm & Extinguishers",
  Flood_Zone: "Low", Cyber_Exposure: "Medium", Hot_Works: "No", Cooking_Deep_Frying: "No",
  Stock_In_Transit: "No", Manual_Work_Away: "No", Prior_Claims_3Y: "None",
  Property_Sum_Insured: 1200000, Manual_Adjustment_Pct: 0, Rating_Class: "C",
};
const REQUIRED_BY_STEP = {
  business: ["Distribution_Channel", "Agent_Tier", "Effective_Date", "Business_Activity", "Annual_Turnover"],
  risk: ["Building_Type", "Construction", "Fire_Protection", "Flood_Zone", "Cyber_Exposure", "Prior_Claims_3Y", "Property_Sum_Insured"],
};

// ─── Mock rating engine (used only if Spark is unreachable) ─────────────────
function simulateUnderwriting(f) {
  const act = ACTIVITY_TABLE[f.Business_Activity] || { appetite: "Accept", trigger: "" };
  const reasons = [];
  let status = "Accept";
  if (act.appetite === "Refer") { status = "Refer"; reasons.push(act.trigger); }
  if (f.Cyber_Exposure === "High") { status = "Refer"; reasons.push("High cyber exposure"); }
  if (f.Flood_Zone === "High") { status = "Refer"; reasons.push("High flood-zone exposure"); }
  if (f.Prior_Claims_3Y === "3+") { status = "Decline"; reasons.push("3+ claims in prior 3 years"); }
  if (Number(f.Property_Sum_Insured) > 5000000) { status = "Refer"; reasons.push("Sum insured exceeds standard authority"); }
  if (f.Construction === "Light / Metal" && f.Fire_Protection === "None") { status = "Decline"; reasons.push("Light/metal construction with no fire protection"); }
  return {
    UW_Status: status,
    UW_Reasons: reasons.length ? reasons.join("; ") + "; " : "",
    UW_Indicative_Actions: status === "Accept" ? "Proceed to quote" : status === "Refer" ? "Refer to UW / Control" : "Decline — do not quote",
    UW_Commercial_Action: status === "Accept" ? "Straight-through quote issued" : "Route to underwriter / Control workflow",
  };
}
function ratingBaseAndLoad(f) {
  const base = Number(f.Property_Sum_Insured || 0) * 0.0021 + Number(f.Annual_Turnover || 0) * 0.0009;
  const riskLoad = 1
    + (f.Fire_Protection === "None" ? 0.22 : f.Fire_Protection === "Basic Extinguishers" ? 0.1 : f.Fire_Protection === "Full Sprinklers" ? -0.08 : 0)
    + (f.Cyber_Exposure === "High" ? 0.15 : f.Cyber_Exposure === "Medium" ? 0.05 : 0)
    + (f.Cooking_Deep_Frying === "Yes" ? 0.08 : 0)
    + (f.Hot_Works === "Yes" ? 0.05 : 0)
    + Number(f.Manual_Adjustment_Pct || 0);
  return { base, riskLoad };
}
function simulateQuickQuote(f) {
  const { base, riskLoad } = ratingBaseAndLoad(f);
  const p2 = Math.max(650, Math.round(base * riskLoad));
  const tiers = [
    { code: "PACRET1", name: "Retail Essential", propLimit: Math.round(Number(f.Property_Sum_Insured) * 0.6), liabLimit: 1000000, bi: 0,       mult: 0.68 },
    { code: "PACRET2", name: "Retail Standard",  propLimit: Number(f.Property_Sum_Insured),                liabLimit: 2000000, bi: 250000, mult: 1.0 },
    { code: "PACRET3", name: "Retail Premier",   propLimit: Math.round(Number(f.Property_Sum_Insured) * 1.4), liabLimit: 5000000, bi: 500000, mult: 1.45 },
  ];
  const rows = [
    { Package: "Package Code", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, t.code])) },
    { Package: "Package Name", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, t.name])) },
    { Package: "Property Limit", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, t.propLimit])) },
    { Package: "Liability Limit", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, t.liabLimit])) },
    { Package: "Business Interruption", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, t.bi])) },
    { Package: "Annual Premium", ...Object.fromEntries(tiers.map((t, i) => [`Offer ${i + 1}`, Math.round(p2 * t.mult)])) },
  ];
  return { "Quick_Quote.Offer_Count": 3, "Quick_Quote.Offers": rows };
}
function simulateQuickBind(f, packageCode) {
  const { "Quick_Quote.Offers": rows } = simulateQuickQuote(f);
  const idx = Object.keys(rows[0]).find(k => rows.find(r => r.Package === "Package Code")[k] === packageCode);
  const get = (label) => rows.find(r => r.Package === label)?.[idx];
  const premium = get("Annual Premium") || 0;
  return {
    "Quick_Quote_Bind.Coverages": [
      { Coverage: "Property", "Coverage Code": "PROP", "Coverage Description": "Property Damage & Fire", Limit: get("Property Limit"), Deductible: 500, "Coverage Premium": Math.round(premium * 0.55) },
      { Coverage: "Liability", "Coverage Code": "LIAB", "Coverage Description": "Public & Products Liability", Limit: get("Liability Limit"), Deductible: null, "Coverage Premium": Math.round(premium * 0.3) },
      { Coverage: "Business Interruption", "Coverage Code": "BI", "Coverage Description": "Business Interruption", Limit: get("Business Interruption"), Deductible: null, "Coverage Premium": Math.round(premium * 0.15) },
    ],
    "Quick_Quote_Bind.Premium": premium,
  };
}
// Real payload evidence (from a captured Coherent tester request) shows Flex_Quote_Bind
// needs a "Package Code" entry AND 6 coverage keys, not the 3 I originally modeled:
// Property, Liability, Business Interruption, Employers Liability, Cyber, Endorsement 1.
// Also, the real Quick_Quote.Offers data shows Property/Liability/Employers Liability are
// "Required" on every tier (i.e. always included, never optional) — only BI/Cyber/END1 vary.
// Confirmed segment-specific: v2's Flex_Quote_List returns "PACEDUC" as the Package
// Code for Education & Childcare (vs the old "PACRETC" guess for Retail) — so this is
// read live via parseFlexPackageCode() below. This constant is only a last-resort
// fallback if that live read fails for some reason.
const FLEX_PACKAGE_CODE_FALLBACK = "PACRETC";
const FLEX_MANDATORY = ["Property", "Liability", "Employers Liability"];
const FLEX_DESCRIPTIONS = {
  Property: "Property Damage & Fire", Liability: "Public & Products Liability",
  "Employers Liability": "Employers Liability", "Business Interruption": "Business Interruption",
  Cyber: "Cyber Liability", "Endorsement 1": "Food Spoilage Coverage",
};
// Flex config is On/Off only per the real Flex_Quote_Bind_Coverages shape — each
// coverage always sends its standard-tier code, the front-end doesn't offer a tier picker.
const FLEX_FIXED_CODE = {
  Property: "PROP2", Liability: "LIAB2", "Employers Liability": "EL2",
  "Business Interruption": "BI2", Cyber: "CYBER2", "Endorsement 1": "FSC1",
};
function simulateFlexBind(f, selections) {
  const base = Number(f.Property_Sum_Insured || 0) * 0.0021 + Number(f.Annual_Turnover || 0) * 0.0009;
  const shareByCov = { Property: 0.4, Liability: 0.22, "Employers Liability": 0.18, "Business Interruption": 0.12, Cyber: 0.15, "Endorsement 1": 0.03 };
  const limitBaseByCov = { Property: Number(f.Property_Sum_Insured), Liability: 2000000, "Employers Liability": 1000000, "Business Interruption": 250000, Cyber: 500000 };
  const rows = [];
  Object.entries(selections).forEach(([cov, include]) => {
    if (!include) return;
    const code = FLEX_FIXED_CODE[cov];
    if (cov === "Endorsement 1") {
      rows.push({ Coverage: cov, "Coverage Code": code, "Coverage Description": FLEX_DESCRIPTIONS[cov], Limit: null, Deductible: null, "Coverage Premium": 75 });
      return;
    }
    rows.push({
      Coverage: cov, "Coverage Code": code, "Coverage Description": FLEX_DESCRIPTIONS[cov],
      Limit: Math.round((limitBaseByCov[cov] || 0) * 1.0), Deductible: cov === "Property" ? 500 : null,
      "Coverage Premium": Math.round(base * (shareByCov[cov] || 0.1) * 3),
    });
  });
  const premium = rows.reduce((s, r) => s + (Number(r["Coverage Premium"]) || 0), 0);
  return { "Flex_Quote_Bind.Coverages": rows, "Flex_Quote_Bind.Coverages_Valid": rows.length ? "Valid" : "No coverages selected", "Flex_Quote_Bind.Premium": premium };
}

// ─── Formatting & small utilities ───────────────────────────────────────────
function fmtMoney(v, ccy = "USD") {
  const n = Number(v);
  if (v === null || v === undefined || v === "" || isNaN(n)) return "—";
  return n.toLocaleString("en-US", { style: "currency", currency: ccy, maximumFractionDigits: 0 });
}
function fmtNum(v) {
  const n = Number(v);
  if (v === null || v === undefined || v === "" || isNaN(n)) return "—";
  return n.toLocaleString("en-US");
}
function hashCode(s) { let h = 0; for (let i = 0; i < s.length; i++) { h = (h << 5) - h + s.charCodeAt(i); h |= 0; } return h; }

// ─── API log CSV export ─────────────────────────────────────────────────────
function csvEscape(v) {
  const s = String(v ?? "");
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}
function apiLogToCSV(log) {
  const headers = ["Timestamp", "Type", "Purpose", "Service Category", "Status", "HTTP Status", "Duration (ms)", "Request", "Response"];
  const rows = log.map(e => [e.timestamp, e.type, e.purpose, e.serviceCategory, e.status, e.httpStatus ?? "", e.durationMs, e.requestBody, e.responseBody]);
  return [headers, ...rows].map(r => r.map(csvEscape).join(",")).join("\n");
}
function downloadApiLogCSV(log) {
  const csv = apiLogToCSV(log);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `sme-product-quotation-api-log-${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function useCountUp(target, duration = 650) {
  const [val, setVal] = useState(target || 0);
  const prev = useRef(target || 0);
  useEffect(() => {
    const from = prev.current, to = Number(target) || 0;
    if (from === to) return;
    const start = performance.now();
    let raf;
    function tick(now) {
      const p = Math.min(1, (now - start) / duration);
      const eased = 1 - Math.pow(1 - p, 3);
      setVal(from + (to - from) * eased);
      if (p < 1) raf = requestAnimationFrame(tick);
      else prev.current = to;
    }
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [target, duration]);
  return val;
}

// ─── Schema helpers (validation-endpoint driven) ────────────────────────────
function schemaOf(schema, field) { return (schema && schema[field]) || {}; }
function fieldOptions(schema, field, fallback) { return schemaOf(schema, field).options || fallback || []; }
function fieldRangeError(schema, field, value) {
  const s = schemaOf(schema, field);
  const n = Number(value);
  if (s.min != null && value !== "" && !isNaN(n) && n < s.min) return s.input_message || `Minimum ${s.min}`;
  if (s.max != null && value !== "" && !isNaN(n) && n > s.max) return s.input_message || `Maximum ${s.max}`;
  return null;
}

// ─── UI primitives (matched to ALHRS card/shadow/input conventions) ────────
function Field({ label, hint, error, children, span }) {
  return (
    <div style={{ gridColumn: span ? `span ${span}` : undefined, marginBottom: 16 }}>
      <div style={{ fontSize: 11, fontWeight: 600, color: T.muted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5, fontFamily: FONT }}>{label}</div>
      {children}
      {error
        ? <div style={{ fontSize: 11, color: T.error, marginTop: 4, fontWeight: 500 }}>⚠ {error}</div>
        : hint && <div style={{ fontSize: 11, color: T.muted, marginTop: 4, lineHeight: 1.4 }}>{hint}</div>}
    </div>
  );
}
function inputBase(hasError) {
  return {
    width: "100%", height: 34, boxSizing: "border-box", fontFamily: FONT, fontSize: 13,
    padding: "0 10px", borderRadius: 7, border: `1px solid ${hasError ? T.error : T.border}`,
    background: T.inputBg, color: T.text, outline: "none", lineHeight: "32px",
    boxShadow: hasError ? `0 0 0 2px ${T.errorLight}` : "none", transition: "border 0.15s, box-shadow 0.15s",
  };
}
function Select({ value, onChange, options, disabled, error }) {
  return (
    <select value={value} disabled={disabled} onChange={e => onChange(e.target.value)}
      style={{ ...inputBase(error), cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.6 : 1, appearance: "none",
        backgroundImage: `url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24'%3E%3Cpath fill='%236B6B8A' d='M7 10l5 5 5-5z'/%3E%3C/svg%3E")`,
        backgroundRepeat: "no-repeat", backgroundPosition: "right 10px center" }}
      onFocus={e => { e.target.style.borderColor = T.navy; e.target.style.boxShadow = `0 0 0 2px rgba(14,27,61,0.12)`; }}
      onBlur={e => { e.target.style.borderColor = error ? T.error : T.border; e.target.style.boxShadow = error ? `0 0 0 2px ${T.errorLight}` : "none"; }}>
      {options.map(o => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}
function TextInput({ value, onChange, type = "text", prefix, disabled, error }) {
  return (
    <div style={{ position: "relative", display: "flex", alignItems: "center" }}>
      {prefix && <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", fontSize: 13, color: T.muted, lineHeight: 1 }}>{prefix}</span>}
      <input type={type} value={value} disabled={disabled}
        onChange={e => onChange(type === "number" ? e.target.value.replace(/[^\d.]/g, "") : e.target.value)}
        style={{ ...inputBase(error), paddingLeft: prefix ? 22 : 10, opacity: disabled ? 0.6 : 1, width: "100%" }}
        onFocus={e => { e.target.style.borderColor = T.navy; e.target.style.boxShadow = `0 0 0 2px rgba(14,27,61,0.12)`; }}
        onBlur={e => { e.target.style.borderColor = error ? T.error : T.border; e.target.style.boxShadow = error ? `0 0 0 2px ${T.errorLight}` : "none"; }} />
    </div>
  );
}
function Toggle({ value, onChange }) {
  const on = value === "Yes";
  return (
    <button onClick={() => onChange(on ? "No" : "Yes")} style={{ display: "flex", alignItems: "center", gap: 10, background: "none", border: "none", cursor: "pointer", padding: 0, lineHeight: 1 }}>
      <div style={{ width: 38, height: 21, borderRadius: 12, background: on ? T.navy : T.mid, position: "relative", transition: "background 0.15s", flexShrink: 0 }}>
        <div style={{ width: 15, height: 15, borderRadius: "50%", background: "#fff", position: "absolute", top: 3, left: on ? 20 : 3, transition: "left 0.15s", boxShadow: "0 1px 2px rgba(0,0,0,0.3)" }} />
      </div>
      <span style={{ fontSize: 13, fontWeight: 600, color: on ? T.navy : T.muted, fontFamily: FONT, lineHeight: "21px" }}>{value}</span>
    </button>
  );
}
function ToggleRow({ label, value, onChange }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 0", borderBottom: `1px solid ${T.mid}` }}>
      <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{label}</span>
      <Toggle value={value} onChange={onChange} />
    </div>
  );
}
function Pill({ children, tone = "muted" }) {
  const map = {
    muted: { bg: T.mid, fg: T.muted }, red: { bg: T.errorLight, fg: T.error }, gold: { bg: T.warnLight, fg: T.warn },
    green: { bg: T.successLight, fg: T.success }, navy: { bg: T.navy, fg: "#fff" },
  }[tone];
  return <span style={{ display: "inline-flex", alignItems: "center", padding: "4px 12px", borderRadius: 20, fontSize: 11, fontWeight: 700, letterSpacing: "0.04em", background: map.bg, color: map.fg, fontFamily: FONT, textTransform: "uppercase" }}>{children}</span>;
}
function Card({ children, style, id }) {
  return <div id={id} style={{ background: T.white, borderRadius: 12, border: `1px solid ${T.mid}`, padding: "24px 28px", boxShadow: "0 2px 16px rgba(14,27,61,0.06)", ...style }}>{children}</div>;
}
function SectionTitle({ eyebrow, title, subtitle }) {
  return (
    <div style={{ display: "flex", alignItems: "stretch", gap: 16, marginBottom: 24, textAlign: "left" }}>
      <div style={{ width: 4, background: T.red, borderRadius: 2, flexShrink: 0 }} />
      <div style={{ textAlign: "left", flex: 1 }}>
        {eyebrow && <div style={{ fontSize: 11, fontWeight: 700, letterSpacing: "0.1em", color: T.red, textTransform: "uppercase", marginBottom: 4, textAlign: "left" }}>{eyebrow}</div>}
        <div style={{ fontFamily: FONT, fontSize: 22, fontWeight: 800, color: T.navy, textAlign: "left" }}>{title}</div>
        {subtitle && <div style={{ fontSize: 13, color: T.muted, marginTop: 5, lineHeight: 1.5, maxWidth: 560, textAlign: "left" }}>{subtitle}</div>}
      </div>
    </div>
  );
}
function PrimaryButton({ children, onClick, disabled, tone = "red", full }) {
  const bg = tone === "red" ? T.red : tone === "navy" ? T.navy : T.success;
  return (
    <button onClick={onClick} disabled={disabled} style={{
      background: disabled ? T.muted : bg, color: "#fff", border: "none", borderRadius: 8, padding: "11px 22px",
      fontSize: 13, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer", fontFamily: FONT,
      width: full ? "100%" : undefined, boxShadow: disabled ? "none" : `0 2px 12px rgba(208,38,58,0.28)`, transition: "opacity 0.15s",
    }} onMouseEnter={e => { if (!disabled) e.currentTarget.style.opacity = 0.88; }} onMouseLeave={e => { e.currentTarget.style.opacity = 1; }}>
      {children}
    </button>
  );
}
function GhostButton({ children, onClick, disabled }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      background: T.white, color: T.navy, border: `1px solid ${T.mid}`, borderRadius: 8, padding: "10px 20px",
      fontSize: 13, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer", fontFamily: FONT, opacity: disabled ? 0.5 : 1,
    }}>{children}</button>
  );
}
function Spinner({ light }) {
  return <div style={{ width: 14, height: 14, border: `2px solid ${light ? "rgba(255,255,255,0.35)" : T.mid}`, borderTopColor: light ? "#fff" : T.red, borderRadius: "50%", animation: "spin 0.7s linear infinite" }} />;
}
function Skeleton({ w = "100%", h = 14, style }) {
  return <div style={{ width: w, height: h, borderRadius: 5, background: `linear-gradient(90deg, ${T.mid} 25%, #F0F0F6 37%, ${T.mid} 63%)`, backgroundSize: "400% 100%", animation: "shimmer 1.4s ease infinite", ...style }} />;
}
function FadeIn({ children, k }) { return <div key={k} style={{ animation: "fadeSlide 0.3s ease" }}>{children}</div>; }

// ─── Appetite Bar — flatter, matched to the card/pill language ─────────────
function AppetiteBar({ status }) {
  const posFor = (st) => st === "Decline" ? 16.5 : st === "Refer" ? 50 : st ? 83.5 : 0;
  const [pos, setPos] = useState(0);
  useEffect(() => { const t = setTimeout(() => setPos(posFor(status)), 150); return () => clearTimeout(t); }, [status]);
  const tone = status === "Decline" ? "red" : status === "Refer" ? "gold" : status ? "green" : "muted";
  return (
    <div style={{ maxWidth: 460, margin: "0 auto" }}>
      <div style={{ position: "relative", height: 14, borderRadius: 8, overflow: "visible", display: "flex" }}>
        <div style={{ flex: 1, background: T.error, borderRadius: "8px 0 0 8px" }} />
        <div style={{ flex: 1, background: T.warn }} />
        <div style={{ flex: 1, background: T.success, borderRadius: "0 8px 8px 0" }} />
        <div style={{
          position: "absolute", top: -5, left: `calc(${pos}% - 12px)`, width: 24, height: 24, borderRadius: "50%",
          background: T.navy, border: "3px solid #fff", boxShadow: "0 2px 10px rgba(14,27,61,0.35)",
          transition: "left 0.9s cubic-bezier(.2,.8,.2,1)",
        }} />
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 10, fontSize: 10.5, fontWeight: 700, letterSpacing: "0.06em", fontFamily: FONT }}>
        <span style={{ color: T.error }}>DECLINE</span>
        <span style={{ color: T.warn }}>REFER</span>
        <span style={{ color: T.success }}>ACCEPT</span>
      </div>
      <div style={{ textAlign: "center", marginTop: 18 }}><Pill tone={tone}>{status || "Evaluating"}</Pill></div>
    </div>
  );
}

// ─── Header — brand lockup (logos) + status ─────────────────────────────────
function LogoBox({ src, fallback }) {
  const [failed, setFailed] = useState(false);
  if (failed) {
    return <div style={{ height: 34, padding: "0 12px", background: T.mid, borderRadius: 6, display: "flex", alignItems: "center", fontSize: 11.5, fontWeight: 800, color: T.navy, letterSpacing: "0.04em" }}>{fallback}</div>;
  }
  return <img src={src} alt={fallback} onError={() => setFailed(true)} style={{ height: 34, maxWidth: 96, objectFit: "contain" }} />;
}
function ConnectionBadge({ live, checking }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
      <div style={{ width: 7, height: 7, borderRadius: "50%", background: checking ? T.warn : live ? T.success : T.muted }} />
      <span style={{ fontSize: 11, color: T.muted, fontWeight: 500, whiteSpace: "nowrap" }}>
        {checking ? "Connecting…" : live ? "Live Spark" : "Offline demo"}
      </span>
    </div>
  );
}

// ─── Horizontal StepBar (pill nav, ALHRS-style numbered circles) ───────────
const STEPS = [
  { key: "business", label: "Business" }, { key: "risk", label: "Risk" }, { key: "appetite", label: "Appetite" },
  { key: "packages", label: "Packages" }, { key: "bind", label: "Bind" }, { key: "issue", label: "Issue" },
];
function StepBar({ step, reachable, onJump }) {
  const idx = STEPS.findIndex(s => s.key === step);
  return (
    <div style={{ display: "flex", alignItems: "center", background: T.white, borderRadius: 10, padding: "6px 8px", boxShadow: "0 2px 12px rgba(14,27,61,0.08)" }}>
      {STEPS.map((s, i) => {
        const active = i === idx, done = i < idx, canJump = reachable.includes(s.key);
        return (
          <div key={s.key} style={{ display: "flex", alignItems: "center", flex: 1 }}>
            <button onClick={() => canJump && onJump(s.key)} style={{
              display: "flex", alignItems: "center", gap: 8, flex: 1, background: active ? T.navy : done ? T.successLight : "transparent",
              color: active ? "#fff" : done ? T.success : T.muted, border: "none", borderRadius: 8, padding: "9px 12px",
              cursor: canJump ? "pointer" : "default", fontWeight: active ? 700 : 500, fontSize: 12.5, fontFamily: FONT, transition: "all 0.2s",
            }}>
              <span style={{ width: 21, height: 21, borderRadius: "50%", background: active ? T.red : done ? T.success : T.mid, color: active || done ? "#fff" : T.muted, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10.5, fontWeight: 700, flexShrink: 0 }}>
                {done ? "✓" : i + 1}
              </span>
              {s.label}
            </button>
            {i < STEPS.length - 1 && <div style={{ width: 14, height: 2, background: done ? T.success : T.mid, flexShrink: 0 }} />}
          </div>
        );
      })}
    </div>
  );
}

// ─── Live indicative-pricing ticker (right rail) ────────────────────────────
function LiveTicker({ form, ticker, onCheckNow }) {
  const low = useCountUp(ticker.low ?? 0);
  const high = useCountUp(ticker.high ?? 0);
  const status = ticker.status;
  const tone = status === "Decline" ? "red" : status === "Refer" ? "gold" : status ? "green" : "muted";
  const isLoading = ticker.source === "loading";
  const isLive = ticker.source === "live";
  return (
    <div style={{ width: 250, flexShrink: 0 }}>
      <div style={{ position: "sticky", top: 118 }}>
        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: "0.08em", color: T.muted, textTransform: "uppercase", marginBottom: 10 }}>Live Snapshot</div>
        <Card style={{ padding: "18px 18px", marginBottom: 12 }}>
          <div style={{ fontSize: 11, color: T.muted, marginBottom: 2 }}>Business</div>
          <div style={{ fontFamily: FONT, fontSize: 16, fontWeight: 800, color: T.navy, marginBottom: 12, lineHeight: 1.25 }}>{form.Business_Activity || "—"}</div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: T.muted, marginBottom: 4 }}>
            <span>Turnover</span><span style={{ color: T.text, fontWeight: 700 }}>{fmtMoney(form.Annual_Turnover)}</span>
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: T.muted }}>
            <span>Sum Insured</span><span style={{ color: T.text, fontWeight: 700 }}>{fmtMoney(form.Property_Sum_Insured)}</span>
          </div>
        </Card>
        <Card style={{ padding: "18px 18px", marginBottom: 12 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <div style={{ fontSize: 11, color: T.muted }}>Preliminary signal</div>
            {isLoading ? <Skeleton w={54} h={18} style={{ borderRadius: 20 }} /> : <Pill tone={tone}>{status || "—"}</Pill>}
          </div>
          {isLoading ? <Skeleton h={12} /> : (
            <>
              <div style={{ fontSize: 13, fontWeight: 700, color: T.navy, marginBottom: 4 }}>{ticker.indicativeAction || "—"}</div>
              {ticker.reasons && <div style={{ fontSize: 11.5, color: T.muted, lineHeight: 1.5 }}>{ticker.reasons}</div>}
            </>
          )}
          <div style={{ marginTop: 10 }}>
            <button onClick={onCheckNow} disabled={isLoading} style={{
              width: "100%", padding: "8px 10px", borderRadius: 6, border: `1px solid ${T.mid}`, background: T.white,
              color: T.navy, fontSize: 11.5, fontWeight: 700, cursor: isLoading ? "not-allowed" : "pointer", fontFamily: FONT, opacity: isLoading ? 0.6 : 1,
            }}>
              {isLoading ? "Checking…" : "Check Eligibility Now"}
            </button>
          </div>
          <div style={{ fontSize: 10, color: T.muted, opacity: 0.75, marginTop: 8 }}>Confirmed at the Appetite Check step.</div>
        </Card>
        <Card style={{ padding: "18px 18px", background: T.navy, border: "none" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.6)" }}>Indicative annual premium</div>
            <span style={{ fontSize: 9, fontWeight: 800, letterSpacing: "0.06em", color: isLive ? "#3FCB8A" : "rgba(255,255,255,0.5)" }}>
              {isLoading ? "…" : isLive ? "LIVE · SPARK" : "ESTIMATE (OFFLINE)"}
            </span>
          </div>
          {isLoading && ticker.low == null ? (
            <Skeleton w="80%" h={22} style={{ background: "rgba(255,255,255,0.15)" }} />
          ) : (
            <div style={{ fontSize: 20, fontWeight: 800, color: "#fff" }}>
              {fmtMoney(low)} <span style={{ fontSize: 13, color: "rgba(255,255,255,0.5)", fontWeight: 400 }}>–</span> {fmtMoney(high)}
            </div>
          )}
          <div style={{ fontSize: 10.5, color: "rgba(255,255,255,0.45)", marginTop: 8, lineHeight: 1.5 }}>
            {isLive ? "From Spark's Quick Quote packages, recalculated as you type." : "Spark unreachable — showing an offline approximation."}
          </div>
        </Card>
      </div>
    </div>
  );
}

// ─── "Reasoning" transition overlay between Risk and Appetite ──────────────
function CalculatingOverlay({ visible }) {
  const messages = ["Resolving activity appetite…", "Evaluating risk factors…", "Checking referral thresholds…", "Calling Spark rating engine…"];
  const [i, setI] = useState(0);
  useEffect(() => {
    if (!visible) { setI(0); return; }
    const t = setInterval(() => setI(v => Math.min(messages.length - 1, v + 1)), 420);
    return () => clearInterval(t);
  }, [visible]);
  if (!visible) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(14,27,61,0.92)", zIndex: 600, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ textAlign: "center", maxWidth: 340 }}>
        <div style={{ margin: "0 auto 22px", width: 42, height: 42, border: "3px solid rgba(255,255,255,0.2)", borderTopColor: T.red, borderRadius: "50%", animation: "spin 0.9s linear infinite" }} />
        {messages.map((m, mi) => (
          <div key={m} style={{ fontSize: 14, fontFamily: FONT, marginBottom: 8, color: mi <= i ? "#fff" : "rgba(255,255,255,0.28)", fontWeight: mi === i ? 700 : 400, transition: "color 0.25s" }}>
            {mi < i ? "✓ " : ""}{m}
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── Banners (ALHRS-style colored alert boxes) ──────────────────────────────
function Banner({ tone, title, children }) {
  const map = { error: { bg: T.errorLight, border: T.error, fg: T.error }, warn: { bg: T.warnLight, border: T.warn, fg: T.warn } }[tone];
  return (
    <div style={{ background: map.bg, border: `1px solid ${map.border}`, borderRadius: 10, padding: "14px 18px", marginBottom: 20 }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: map.fg, marginBottom: children ? 6 : 0 }}>{title}</div>
      {children && <div style={{ fontSize: 12.5, color: map.fg, lineHeight: 1.5 }}>{children}</div>}
    </div>
  );
}

// ─── Business Profile step ──────────────────────────────────────────────────
function StepBusiness({ form, setForm, onNext, errors, schema, activityOptions, activityLoading, resolvedActivity, resolvedLoading }) {
  const set = (k) => (v) => setForm(f => ({ ...f, [k]: v }));
  const resolvedBox = (val) => resolvedLoading
    ? <Skeleton h={34} style={{ borderRadius: 7 }} />
    : <div style={{ ...inputBase(false), background: T.mid, fontWeight: 700, fontSize: 12.5 }}>{val || "—"}</div>;
  return (
    <>
      <SectionTitle eyebrow="Step 01" title="Business Profile" subtitle="Distribution, timing and business identification." />
      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 14, textTransform: "uppercase" }}>Quote Setup</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: "0 20px" }}>
          <Field label="Distribution Channel" error={errors.Distribution_Channel}><Select value={form.Distribution_Channel} onChange={set("Distribution_Channel")} options={fieldOptions(schema, "Distribution_Channel", OPTS.Distribution_Channel)} error={errors.Distribution_Channel} /></Field>
          <Field label="Agent Tier" error={errors.Agent_Tier}><Select value={form.Agent_Tier} onChange={set("Agent_Tier")} options={fieldOptions(schema, "Agent_Tier", OPTS.Agent_Tier)} error={errors.Agent_Tier} /></Field>
          <Field label="Effective Date" error={errors.Effective_Date}><TextInput type="date" value={form.Effective_Date} onChange={set("Effective_Date")} error={errors.Effective_Date} /></Field>
          <Field label="Policy Term (months)"><Select value={String(form.Policy_Term_Months)} onChange={v => set("Policy_Term_Months")(Number(v))} options={["12", "24", "36"]} /></Field>
        </div>
      </Card>
      <Card>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 14, textTransform: "uppercase" }}>Business Profile</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 20px" }}>
          <Field label="Industry Segment" span={3} error={errors.Industry_Segment}>
            <Select value={form.Industry_Segment} onChange={set("Industry_Segment")} options={fieldOptions(schema, "Industry_Segment", INDUSTRY_SEGMENTS)} error={errors.Industry_Segment} />
          </Field>
          <Field label="Business Activity" span={3} error={errors.Business_Activity} hint={!activityLoading && !activityOptions.length ? "No activities returned for this segment yet." : undefined}>
            {activityLoading ? <Skeleton h={34} style={{ borderRadius: 7 }} /> : activityOptions.length ? (
              <Select value={form.Business_Activity} onChange={set("Business_Activity")} options={activityOptions} error={errors.Business_Activity} />
            ) : (
              <div style={{ ...inputBase(false), color: T.muted, background: T.mid }}>Not available for this segment yet</div>
            )}
          </Field>

          <Field label="Activity Code">{resolvedBox(resolvedActivity?.code)}</Field>
          <Field label="Rating Class">{resolvedBox(resolvedActivity?.ratingClass)}</Field>
          <Field label="Occupancy Group">{resolvedBox(resolvedActivity?.occupancyGroup)}</Field>

          <Field label="Years in Business"><TextInput type="number" value={form.Years_In_Business} onChange={set("Years_In_Business")} /></Field>
          <Field label="Employee Count"><TextInput type="number" value={form.Employee_Count} onChange={set("Employee_Count")} /></Field>
          <Field label="Premises Count"><TextInput type="number" value={form.Premises_Count} onChange={set("Premises_Count")} /></Field>

          <Field label="Annual Turnover" span={3} error={errors.Annual_Turnover}><TextInput type="number" prefix="$" value={form.Annual_Turnover} onChange={set("Annual_Turnover")} error={errors.Annual_Turnover} /></Field>
        </div>
      </Card>
      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 20 }}>
        <PrimaryButton onClick={onNext}>Next →</PrimaryButton>
      </div>
    </>
  );
}

// ─── Risk Profile step ──────────────────────────────────────────────────────
function StepRisk({ form, setForm, onNext, onBack, loading, errors, schema }) {
  const set = (k) => (v) => setForm(f => ({ ...f, [k]: v }));
  const rangeErr = (k) => errors[k] || fieldRangeError(schema, k, form[k]);
  return (
    <>
      <SectionTitle eyebrow="Step 02" title="Risk Profile" subtitle="Physical risk, protections and loss history — what the appetite and rating engines evaluate next." />
      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 14, textTransform: "uppercase" }}>Premises &amp; Construction</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 20px" }}>
          <Field label="Building Type" error={errors.Building_Type}><Select value={form.Building_Type} onChange={set("Building_Type")} options={fieldOptions(schema, "Building_Type", OPTS.Building_Type)} error={errors.Building_Type} /></Field>
          <Field label="Construction" error={errors.Construction}><Select value={form.Construction} onChange={set("Construction")} options={fieldOptions(schema, "Construction", OPTS.Construction)} error={errors.Construction} /></Field>
          <Field label="Fire Protection" error={errors.Fire_Protection}><Select value={form.Fire_Protection} onChange={set("Fire_Protection")} options={fieldOptions(schema, "Fire_Protection", OPTS.Fire_Protection)} error={errors.Fire_Protection} /></Field>
          <Field label="Flood Zone" error={errors.Flood_Zone}><Select value={form.Flood_Zone} onChange={set("Flood_Zone")} options={fieldOptions(schema, "Flood_Zone", OPTS.Flood_Zone)} error={errors.Flood_Zone} /></Field>
          <Field label="Cyber Exposure" error={errors.Cyber_Exposure}><Select value={form.Cyber_Exposure} onChange={set("Cyber_Exposure")} options={fieldOptions(schema, "Cyber_Exposure", OPTS.Cyber_Exposure)} error={errors.Cyber_Exposure} /></Field>
          <Field label="Prior Claims (3Y)" error={errors.Prior_Claims_3Y}><Select value={form.Prior_Claims_3Y} onChange={set("Prior_Claims_3Y")} options={fieldOptions(schema, "Prior_Claims_3Y", OPTS.Prior_Claims_3Y)} error={errors.Prior_Claims_3Y} /></Field>
        </div>
      </Card>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 14, textTransform: "uppercase" }}>Operational Hazards</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 40px" }}>
          <ToggleRow label="Hot Works Performed" value={form.Hot_Works} onChange={set("Hot_Works")} />
          <ToggleRow label="Cooking / Deep Frying On-Site" value={form.Cooking_Deep_Frying} onChange={set("Cooking_Deep_Frying")} />
          <ToggleRow label="Stock In Transit" value={form.Stock_In_Transit} onChange={set("Stock_In_Transit")} />
          <ToggleRow label="Manual Work Away From Premises" value={form.Manual_Work_Away} onChange={set("Manual_Work_Away")} />
        </div>
      </Card>
      <Card>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 14, textTransform: "uppercase" }}>Sums Insured &amp; Commercial Terms</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 20px" }}>
          <Field label="Property Sum Insured" error={rangeErr("Property_Sum_Insured")}><TextInput type="number" prefix="$" value={form.Property_Sum_Insured} onChange={set("Property_Sum_Insured")} error={rangeErr("Property_Sum_Insured")} /></Field>
          <Field label="Agent Commission %"><TextInput type="number" value={form.Commission_Pct * 100} onChange={v => setForm(f => ({ ...f, Commission_Pct: Number(v) / 100 }))} /></Field>
          <Field label="Manual Pricing Adjustment %" hint="Leave at 0 unless a referral authority has approved a loading/discount.">
            <TextInput type="number" value={form.Manual_Adjustment_Pct * 100} onChange={v => setForm(f => ({ ...f, Manual_Adjustment_Pct: Number(v) / 100 }))} />
          </Field>
        </div>
      </Card>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 20 }}>
        <GhostButton onClick={onBack}>← Previous</GhostButton>
        <PrimaryButton onClick={onNext} disabled={loading}>
          {loading ? <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Spinner light /> Checking appetite…</span> : "Run Appetite Check →"}
        </PrimaryButton>
      </div>
    </>
  );
}

// ─── Appetite step ───────────────────────────────────────────────────────────
function StepAppetite({ uw, onBack, onProceed, onNewQuote }) {
  const status = uw?.UW_Status;
  // Only treat explicit Refer/Decline as a hard stop — anything else (Accept, Auto Quote,
  // or any other accept-equivalent string the workbook may use) is treated as proceedable.
  // This was tightened after finding the real workbook can return "Auto Quote" rather than
  // "Accept" — gating on `status === "Accept"` alone would have wrongly blocked that case.
  const blocking = status === "Refer" || status === "Decline";
  const unresolved = !status;
  const tone = status === "Decline" ? "red" : status === "Refer" ? "gold" : status ? "green" : "muted";
  return (
    <>
      <SectionTitle eyebrow="Step 03" title="Underwriting Appetite Check" subtitle="Every submission is scored against activity appetite, risk factors and referral thresholds — instantly, using the exact logic from the underwriting workbook." />
      <Card style={{ textAlign: "center", padding: "32px 28px" }}>
        <AppetiteBar status={status} />
        <div style={{ fontFamily: FONT, fontSize: 18, fontWeight: 800, color: T.navy, marginTop: 18 }}>{uw?.UW_Indicative_Actions}</div>
        {uw?.UW_Reasons && <div style={{ fontSize: 13, color: T.muted, marginTop: 8, maxWidth: 440, margin: "8px auto 0", lineHeight: 1.5 }}>{uw.UW_Reasons}</div>}
      </Card>
      <Card style={{ marginTop: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 12, textTransform: "uppercase" }}>Underwriting Decision</div>
        <DocRow label="UW Status" value={status || "—"} />
        <DocRow label="UW Indicative Action" value={uw?.UW_Indicative_Actions || "—"} />
        <DocRow label="UW Reasons" value={uw?.UW_Reasons || "—"} />
        <DocRow label="UW Commercial Action" value={uw?.UW_Commercial_Action || "—"} />
      </Card>
      {blocking && (
        <div style={{ marginTop: 16 }}>
          <Banner tone={tone === "red" ? "error" : "warn"} title={status === "Decline" ? "⛔ This risk falls outside appetite" : "⚠ This case requires underwriter review"}>
            {uw?.UW_Commercial_Action}. {status === "Refer"
              ? "The submission has been routed to Control for underwriter review and approval routing. No indicative price can be shown to the agent until a decision is recorded."
              : "The submission cannot proceed to quote. The agent has been notified and the case is closed on this channel."}
            <div style={{ fontSize: 11, opacity: 0.8, marginTop: 8, fontWeight: 600 }}>CONTROL-REF: CTRL-{Math.abs(hashCode(JSON.stringify(uw))).toString().slice(0, 8)}</div>
          </Banner>
        </div>
      )}
      {unresolved && (
        <div style={{ marginTop: 16 }}>
          <Banner tone="warn" title="⚠ Underwriting status could not be determined">
            The Spark response didn't include a UW_Status — this usually means Activity_Code/Activity_Appetite/Activity_UW_Trigger didn't resolve for this Business Activity. Check the Activity Mapping field on the Business Profile step.
          </Banner>
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 20 }}>
        <GhostButton onClick={onBack}>← Previous</GhostButton>
        {!blocking && !unresolved ? <PrimaryButton onClick={onProceed}>View Package Options →</PrimaryButton> : <PrimaryButton tone="navy" onClick={onNewQuote}>Start New Quote</PrimaryButton>}
      </div>
    </>
  );
}

// ─── Package comparison (Quick Quote) ───────────────────────────────────────
function humanize(label) {
  return String(label ?? "").replace(/_/g, " ").trim();
}
function escapeRegex(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
// Spark abbreviates some coverage names in row labels (EL, BI, END1) but not others —
// mapping to consistent, user-friendly names used everywhere (package cards, Flex
// builder, requirements lookup) so the same coverage is never shown two different ways.
const COVERAGE_DISPLAY_NAMES = { EL: "Employers Liability", BI: "Business Interruption", END1: "Endorsement 1" };
function friendlyCoverageName(rawBase) {
  return COVERAGE_DISPLAY_NAMES[rawBase] || humanize(rawBase);
}
// Parses xOutput_Industry_and_Business_Mapping into lookup tables in both directions.
// Shape is unconfirmed from a live capture yet — this assumes the most natural flat
// shape (rows of {Industry_Segment, Business_Activity}). Flag if this doesn't parse:
// send a sample response and I'll adjust to match the real shape.
function parseIndustryMapping(list) {
  if (!Array.isArray(list) || !list.length) return null;
  const first = list[0];
  if (first && typeof first === "object" && "Industry_Segment" in first && "Business_Activity" in first) {
    const bySegment = {}, byActivity = {};
    list.forEach(r => {
      const seg = r.Industry_Segment, act = r.Business_Activity;
      if (!seg || !act) return;
      if (!bySegment[seg]) bySegment[seg] = [];
      bySegment[seg].push(act);
      byActivity[act] = seg;
    });
    return { bySegment, byActivity };
  }
  return null; // shape didn't match what we expected — caller falls back to per-segment fetching
}
// layout to Quick_Quote.Offers — a label column (e.g. "Package Compose") plus
// "Offer 1"/"Offer 2"/"Offer 3" columns, where only as many columns as "Max Options"
// actually contain data (the rest come back blank/#N/A). We find the "_include" rows
// and read Required/Optional/None from whichever offer column is actually populated.
function parseFlexRequirements(list) {
  if (!Array.isArray(list) || !list.length) return null;
  const labelKey = Object.keys(list[0])[0];
  const offerKeys = Object.keys(list[0]).filter(k => k !== labelKey);
  const nameRow = list.find(r => /name/i.test(String(r[labelKey])));
  const validKey = offerKeys.find(k => {
    const v = nameRow?.[k];
    return v != null && v !== "" && !/^#N\/A$/i.test(String(v));
  }) || offerKeys[0];
  if (!validKey) return null;
  const reqs = {};
  list.forEach(row => {
    const label = String(row[labelKey] ?? "");
    // Use the same friendly-name mapping as groupCoverageAttrs, so keys here (e.g.
    // "Employers Liability") match what FlexBuilder actually looks up — previously this
    // stored the raw abbreviation ("EL") which never matched, silently breaking the
    // Required/Optional lookup for exactly the three abbreviated coverages.
    if (/_include$/i.test(label)) reqs[friendlyCoverageName(label.replace(/_include$/i, ""))] = row[validKey];
  });
  return Object.keys(reqs).length ? reqs : null;
}
// Reads the actual "Package Code" row from Flex_Quote_List — confirmed segment-specific
// (e.g. "PACEDUC" for Education vs "PACRETC" for Retail), so this must be read live per
// segment rather than hardcoded.
function parseFlexPackageCode(list) {
  if (!Array.isArray(list) || !list.length) return null;
  const labelKey = Object.keys(list[0])[0];
  const offerKeys = Object.keys(list[0]).filter(k => k !== labelKey);
  const codeRow = list.find(r => /^package code$/i.test(String(r[labelKey])));
  if (!codeRow) return null;
  const validKey = offerKeys.find(k => {
    const v = codeRow[k];
    return v != null && v !== "" && !/^#N\/A$/i.test(String(v));
  });
  return validKey ? codeRow[validKey] : null;
}
// The raw sheet rows come as separate "X_include" / "X Coverage" / "X Deductible" rows
// per coverage. Grouping them into one clean block per coverage (rather than showing
// three raw rows each) is both easier to read and removes the underscore-heavy labels.
function groupCoverageAttrs(attrs) {
  const groups = {};
  const order = [];
  attrs.forEach(({ label, value }) => {
    const l = String(label ?? "");
    let rawBase, kind;
    if (/_include$/i.test(l)) { rawBase = l.replace(/_include$/i, ""); kind = "include"; }
    else if (/deductible$/i.test(l)) { rawBase = l.replace(/\s*deductible$/i, ""); kind = "deductible"; }
    else if (/coverage$/i.test(l)) { rawBase = l.replace(/\s*coverage$/i, ""); kind = "tier"; }
    else { rawBase = l; kind = "other"; }
    rawBase = rawBase.trim();
    if (!groups[rawBase]) { groups[rawBase] = { name: friendlyCoverageName(rawBase), rawBase }; order.push(rawBase); }
    groups[rawBase][kind] = value;
  });
  return order.map(k => groups[k]);
}
function pivotOffers(rows) {
  if (!rows || !rows.length) return [];
  const labelKey = Object.keys(rows[0])[0];
  const offerKeys = Object.keys(rows[0]).filter(k => k !== labelKey);
  // The true total is the table's LAST row ("Premium Required" in the real workbook) —
  // rows above it (Property Premium, Liability Premium, etc.) are component premiums,
  // not the total, so we must not grab the first row that merely contains "premium".
  const totalRow = rows.find(r => /premium required/i.test(String(r[labelKey]))) || rows[rows.length - 1];
  return offerKeys.map(ok => {
    const attrs = rows.filter(r => r !== totalRow && !/name|code|premium/i.test(String(r[labelKey]))).map(r => ({ label: r[labelKey], value: r[ok] }));
    return {
      name: rows.find(r => /package name/i.test(String(r[labelKey])))?.[ok] || rows.find(r => /name/i.test(String(r[labelKey])))?.[ok] || ok,
      code: rows.find(r => /package code/i.test(String(r[labelKey])))?.[ok] || rows.find(r => /code/i.test(String(r[labelKey])))?.[ok] || ok,
      premium: totalRow?.[ok],
      groups: groupCoverageAttrs(attrs),
    };
  }).filter(pkg => !/^#N\/A$/i.test(String(pkg.name))); // hide packages the sheet marks as not eligible for this business
}
function PackageCardSkeleton() {
  return (
    <div style={{ flex: 1, background: T.white, border: `1px solid ${T.mid}`, borderRadius: 12, padding: "22px 20px" }}>
      <Skeleton w="60%" h={16} /><div style={{ height: 8 }} /><Skeleton w="30%" h={10} />
      <div style={{ height: 20 }} /><Skeleton w="70%" h={26} /><div style={{ height: 18 }} />
      {[1, 2, 3].map(i => <div key={i} style={{ marginBottom: 10 }}><Skeleton h={12} /></div>)}
    </div>
  );
}
function PackageCard({ pkg, selected, onSelect }) {
  const p = useCountUp(Number(pkg.premium) || 0);
  return (
    <div onClick={onSelect} style={{
      background: T.white, border: `2px solid ${selected ? T.red : T.mid}`, borderRadius: 12, padding: "22px 20px",
      cursor: "pointer", position: "relative", flex: 1, boxShadow: selected ? "0 6px 20px rgba(208,38,58,0.14)" : "0 2px 12px rgba(14,27,61,0.05)", transition: "all 0.15s",
      display: "flex", flexDirection: "column",
    }}>
      <div style={{ height: 26, marginBottom: 12, display: "flex", alignItems: "center" }}>
        {selected && <Pill tone="red">Selected</Pill>}
      </div>
      <div style={{ fontFamily: FONT, fontSize: 17, fontWeight: 800, color: T.navy, lineHeight: 1.25, minHeight: 42, display: "flex", alignItems: "flex-start" }}>{pkg.name}</div>
      <div style={{ fontSize: 11, color: T.muted, marginTop: 2, fontWeight: 600 }}>{pkg.code}</div>
      <div style={{ fontSize: 28, fontWeight: 800, color: T.text, marginTop: 14 }}>{fmtMoney(p)}</div>
      <div style={{ fontSize: 11.5, color: T.muted, marginBottom: 14 }}>per year</div>
      <div style={{ borderTop: `1px solid ${T.mid}`, paddingTop: 4, flex: 1 }}>
        {pkg.groups.map((g, i) => {
          const notIncluded = g.include === "None" || g.include === "No" || (!g.include && !g.tier);
          const tierLabel = (() => {
            if (!g.tier) return null;
            const raw = humanize(g.tier);
            const prefixes = [g.rawBase, g.name].filter(Boolean).map(escapeRegex);
            const stripped = raw.replace(new RegExp("^(" + prefixes.join("|") + ")\\s*", "i"), "").trim();
            return stripped || raw;
          })();
          return (
            <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr auto", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: i < pkg.groups.length - 1 ? `1px solid ${T.mid}` : "none" }}>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: T.text }}>{g.name}</div>
              <div style={{ textAlign: "right" }}>
                {notIncluded ? (
                  <span style={{ fontSize: 11.5, color: T.muted }}>Not included</span>
                ) : (
                  <>
                    <div style={{ fontSize: 12.5, fontWeight: 700, color: T.navy }}>{tierLabel || (g.include === "Required" ? "Included" : humanize(g.include))}</div>
                    {g.deductible != null && <div style={{ fontSize: 10.5, color: T.muted }}>Deductible {fmtMoney(g.deductible)}</div>}
                  </>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
function StepPackages({ mode, setMode, offers, selected, setSelected, flexState, onBack, onContinue, loadingFlex, loadingOffers, onFlexRecalc, onOpenFlex, flexRequirements, flexReqLoading }) {
  return (
    <>
      <SectionTitle eyebrow="Step 04" title="Package Selection" subtitle="Three ready-made packages, or switch to Flex to build coverage-by-coverage." />
      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        {[["quick", "Standard Packages"], ["flex", "Flex Builder"]].map(([k, label]) => (
          <button key={k} onClick={() => { setMode(k); if (k === "flex") onOpenFlex(); }} style={{
            padding: "9px 18px", borderRadius: 8, border: `1px solid ${mode === k ? T.navy : T.mid}`,
            background: mode === k ? T.navy : T.white, color: mode === k ? "#fff" : T.muted,
            fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: FONT,
          }}>{label}</button>
        ))}
      </div>
      {mode === "quick" && (
        <>
          <div style={{ display: "flex", gap: 16 }}>
            {loadingOffers
              ? [1, 2, 3].map(i => <PackageCardSkeleton key={i} />)
              : offers.map((pkg, i) => <PackageCard key={i} pkg={pkg} selected={selected?.code === pkg.code} onSelect={() => setSelected(pkg)} />)}
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 24 }}>
            <GhostButton onClick={onBack}>← Previous</GhostButton>
            <PrimaryButton onClick={() => onContinue("quick")} disabled={!selected}>Continue to Bind →</PrimaryButton>
          </div>
        </>
      )}
      {mode === "flex" && (
        <FlexBuilder state={flexState} loading={loadingFlex} onRecalc={onFlexRecalc} onBack={onBack} onContinue={onContinue}
          requirements={flexRequirements} requirementsLoading={flexReqLoading} />
      )}
    </>
  );
}

// ─── Flex builder ────────────────────────────────────────────────────────────
function FlexBuilder({ state, loading, onRecalc, onBack, onContinue, requirements, requirementsLoading }) {
  const { selections, setSelections, priced } = state;
  const covs = Object.keys(FLEX_FIXED_CODE);
  const total = priced?.["Flex_Quote_Bind.Premium"];
  const animTotal = useCountUp(Number(total) || 0);
  const reqFor = (cov) => requirements?.[cov] || (FLEX_MANDATORY.includes(cov) ? "Required" : "Optional");
  return (
    <>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 4, textTransform: "uppercase" }}>Coverages</div>
        <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 16 }}>
          {requirementsLoading ? "Checking which coverages apply to this business…" : "Required coverages are always included. Toggle the optional ones on or off, then recalculate."}
        </div>
        {requirementsLoading ? [1, 2, 3, 4, 5, 6].map(i => <div key={i} style={{ padding: "13px 0", borderBottom: `1px solid ${T.mid}` }}><Skeleton h={20} /></div>) : covs.map(cov => {
          const include = selections[cov];
          const req = reqFor(cov);
          const mandatory = req === "Required";
          const notApplicable = req === "None";
          return (
            <div key={cov} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "13px 0", borderBottom: `1px solid ${T.mid}`, opacity: notApplicable ? 0.5 : 1 }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{cov}</div>
              {mandatory ? <Pill tone="navy">Always included</Pill>
                : notApplicable ? <Pill tone="muted">Not applicable</Pill>
                : <Toggle value={include ? "Yes" : "No"} onChange={v => setSelections(s => ({ ...s, [cov]: v === "Yes" }))} />}
            </div>
          );
        })}
        <div style={{ marginTop: 20, display: "flex", justifyContent: "center" }}>
          <PrimaryButton onClick={onRecalc} disabled={loading}>
            {loading ? <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Spinner light /> Repricing…</span> : "Reprice Flex Quote →"}
          </PrimaryButton>
        </div>
      </Card>
      {priced && (
        <Card style={{ marginBottom: 16 }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 12, textTransform: "uppercase" }}>Priced Coverages</div>
          <CoverageTable rows={priced["Flex_Quote_Bind.Coverages"] || []} />
          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 14, borderTop: `1px solid ${T.mid}`, paddingTop: 14 }}>
            <div style={{ fontSize: 20, fontWeight: 800, color: T.navy }}>{fmtMoney(animTotal)} <span style={{ fontSize: 12, color: T.muted, fontWeight: 400 }}>/ year</span></div>
          </div>
        </Card>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 20 }}>
        <GhostButton onClick={onBack}>← Previous</GhostButton>
        <PrimaryButton onClick={() => onContinue("flex")} disabled={!priced || !(priced["Flex_Quote_Bind.Coverages"] || []).length}>Continue to Bind →</PrimaryButton>
      </div>
    </>
  );
}
function cleanCell(v) {
  if (v === null || v === undefined || v === "" || /^#N\/A$/i.test(String(v))) return "—";
  return typeof v === "string" ? humanize(v) : v;
}
// Canonical display order (matches the Flex Builder toggle list) — Spark doesn't
// necessarily return rows in this order, so we reorder rather than showing whatever
// order the API happens to return, which made the two lists hard to compare.
const COVERAGE_ORDER = ["Property", "Liability", "Employers Liability", "Business Interruption", "Cyber", "Endorsement 1"];
function sortCoverageRows(rows) {
  return [...rows].sort((a, b) => {
    const ia = COVERAGE_ORDER.indexOf(a.Coverage), ib = COVERAGE_ORDER.indexOf(b.Coverage);
    return (ia === -1 ? 999 : ia) - (ib === -1 ? 999 : ib);
  });
}
// Defensive: try a couple of plausible key names in case Spark's Flex response uses a
// different premium key than the confirmed Quick_Quote_Bind shape ("Coverage Premium").
function coveragePremium(r) {
  const v = r["Coverage Premium"] ?? r["Premium"] ?? r["CoveragePremium"] ?? r["Coverage_Premium"];
  return v != null && v !== "" ? v : null;
}
function CoverageTable({ rows }) {
  const th = { padding: "9px 12px", background: T.navy, color: "#fff", fontSize: 10.5, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", textAlign: "left" };
  const td = { padding: "10px 12px", fontSize: 13, borderBottom: `1px solid ${T.mid}` };
  const sorted = sortCoverageRows(rows);
  return (
    <table style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead><tr>{["Coverage", "Code", "Description", "Limit", "Deductible", "Premium"].map(h => <th key={h} style={th}>{h}</th>)}</tr></thead>
      <tbody>
        {sorted.map((r, i) => {
          const premium = coveragePremium(r);
          return (
            <tr key={i}>
              <td style={{ ...td, fontWeight: 700 }}>{cleanCell(r.Coverage)}</td>
              <td style={{ ...td, color: T.muted, fontWeight: 600 }}>{cleanCell(r["Coverage Code"])}</td>
              <td style={td}>{cleanCell(r["Coverage Description"])}</td>
              <td style={{ ...td, fontWeight: 600 }}>{r.Limit != null && r.Limit !== "" ? fmtNum(r.Limit) : "—"}</td>
              <td style={{ ...td, fontWeight: 600 }}>{r.Deductible != null && r.Deductible !== "" ? fmtMoney(r.Deductible) : "—"}</td>
              <td style={{ ...td, fontWeight: 700 }}>{premium != null ? fmtMoney(premium) : "—"}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

// ─── Bind step ───────────────────────────────────────────────────────────────
function isBindEffectivelyEmpty(coverages, premiumRaw) {
  const noPremium = !premiumRaw || Number(premiumRaw) === 0;
  const noCodes = coverages.length > 0 && coverages.every(c => !c["Coverage Code"] || c["Coverage Code"] === "None");
  return noPremium && noCodes;
}
function StepBind({ bindResult, bindIssues, source, form, selectedName, selectedPkg, onBack, onContinue, endorsement, setEndorsement, loading, onRun }) {
  const coverages = bindResult?.[source === "flex" ? "Flex_Quote_Bind.Coverages" : "Quick_Quote_Bind.Coverages"] || [];
  const premiumRaw = source === "flex" ? bindResult?.["Flex_Quote_Bind.Premium"] : bindResult?.["Quick_Quote_Bind.Premium"];
  const premium = useCountUp(Number(premiumRaw) || 0);
  const commission = premiumRaw ? Number(premiumRaw) * Number(form.Commission_Pct || 0) : null;
  const end1Status = selectedPkg?.groups?.find(g => g.name === "Endorsement 1")?.include;
  const sparkErrors = bindIssues?.errors || [];
  const looksEmpty = isBindEffectivelyEmpty(coverages, premiumRaw);
  return (
    <>
      <SectionTitle eyebrow="Step 05" title="Bind Quote" subtitle={`Binding ${selectedName || "selected package"} for ${form.Business_Activity}.`} />
      {(sparkErrors.length > 0 || looksEmpty) && (
        <Banner tone="error" title="⚠ Spark returned a $0 bind — this is a validation issue, not a real quote">
          {sparkErrors.length > 0 ? (
            <>
              Spark rejected part of the request even though the call itself succeeded:
              <ul style={{ margin: "6px 0 0", paddingLeft: 18 }}>
                {sparkErrors.map((e, i) => (
                  <li key={i} style={{ marginBottom: 3 }}>
                    <strong>{e.source_path || e.error_type}</strong>: {e.message}
                  </li>
                ))}
              </ul>
            </>
          ) : (
            <>Every coverage came back with no code and a $0 premium, which usually means <strong>Quick_Quote_Bind.Package_Code</strong> (or END1) was rejected by Spark's dropdown validation on the workbook side.</>
          )}
          <div style={{ marginTop: 6 }}>This looks like a data-validation config issue on the Spark service (the Package Code / END1 dropdown may only accept a placeholder value) rather than something fixable from this front-end — worth checking directly in the Coherent tester with the same package code.</div>
        </Banner>
      )}
      {source === "quick" && end1Status !== "None" && (
        <Card style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
            <div>
              <div style={{ fontWeight: 700, fontSize: 14, color: T.text }}>Include standard endorsement wording (END1)</div>
              <div style={{ fontSize: 12, color: T.muted, marginTop: 2 }}>Applies the default policy endorsement pack for this package tier.</div>
            </div>
            {end1Status === "Required" ? <Pill tone="navy">Always included</Pill> : <Toggle value={endorsement ? "Yes" : "No"} onChange={v => setEndorsement(v === "Yes")} />}
          </div>
          <GhostButton onClick={() => onRun()} disabled={loading}>
            {loading ? <span style={{ display: "flex", alignItems: "center", gap: 8 }}><Spinner /> Binding…</span> : bindResult ? "Re-bind with current settings" : "Bind Package"}
          </GhostButton>
        </Card>
      )}
      {coverages.length > 0 ? (
        <Card>
          <div style={{ fontSize: 12, fontWeight: 700, color: T.navy, letterSpacing: "0.05em", marginBottom: 12, textTransform: "uppercase" }}>Coverage Breakdown</div>
          <CoverageTable rows={coverages} />
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 32, marginTop: 14, borderTop: `1px solid ${T.mid}`, paddingTop: 14 }}>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 11, color: T.muted, letterSpacing: "0.05em" }}>AGENT COMMISSION</div>
              <div style={{ fontSize: 15, fontWeight: 700, color: T.text, marginTop: 2 }}>{fmtMoney(commission)}</div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div style={{ fontSize: 11, color: T.muted, letterSpacing: "0.05em" }}>ANNUAL PREMIUM</div>
              <div style={{ fontSize: 22, fontWeight: 800, color: T.navy, marginTop: 2 }}>{fmtMoney(premium)}</div>
            </div>
          </div>
        </Card>
      ) : source === "quick" ? (
        <Card style={{ textAlign: "center", padding: "40px 0", color: T.muted, fontSize: 13 }}>Bind the package to see the coverage breakdown.</Card>
      ) : null}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 20 }}>
        <GhostButton onClick={onBack}>← Previous</GhostButton>
        <PrimaryButton onClick={() => onContinue()} disabled={!coverages.length}>Continue to Issuance →</PrimaryButton>
      </div>
    </>
  );
}

// ─── Issuance step ───────────────────────────────────────────────────────────
function StepIssue({ form, uw, bindResult, source, selectedName, quoteRef, onNewQuote }) {
  const [showDoc, setShowDoc] = useState(false);
  const coverages = bindResult?.[source === "flex" ? "Flex_Quote_Bind.Coverages" : "Quick_Quote_Bind.Coverages"] || [];
  const premiumRaw = source === "flex" ? bindResult?.["Flex_Quote_Bind.Premium"] : bindResult?.["Quick_Quote_Bind.Premium"];
  const premium = useCountUp(Number(premiumRaw) || 0, 900);
  return (
    <>
      <SectionTitle eyebrow="Step 06" title="Policy Issuance" subtitle="Final review before generating the policy schedule for the client." />
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 30px" }}>
          <Field label="Business Activity"><MutedBox>{form.Business_Activity}</MutedBox></Field>
          <Field label="Effective Date"><MutedBox>{form.Effective_Date}</MutedBox></Field>
          <Field label="Package"><MutedBox>{selectedName}</MutedBox></Field>
          <Field label="Underwriting Status"><div style={{ display: "flex", alignItems: "center", height: 34 }}><Pill tone={uw?.UW_Status === "Decline" ? "red" : uw?.UW_Status === "Refer" ? "gold" : "green"}>{uw?.UW_Status || "—"}</Pill></div></Field>
        </div>
      </Card>
      <Card style={{ marginBottom: 20, textAlign: "center", padding: "28px 0" }}>
        <div style={{ fontSize: 11, color: T.muted, letterSpacing: "0.06em" }}>TOTAL ANNUAL PREMIUM</div>
        <div style={{ fontSize: 34, fontWeight: 800, color: T.navy, marginTop: 6 }}>{fmtMoney(premium)}</div>
      </Card>
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <GhostButton onClick={onNewQuote}>Start New Quote</GhostButton>
        <PrimaryButton onClick={() => setShowDoc(true)}>Generate Policy Schedule →</PrimaryButton>
      </div>
      {showDoc && <PolicyDocModal form={form} uw={uw} coverages={coverages} premium={premiumRaw} quoteRef={quoteRef} selectedName={selectedName} onClose={() => setShowDoc(false)} />}
    </>
  );
}
function MutedBox({ children }) { return <div style={{ ...inputBase(false), background: T.mid, display: "flex", alignItems: "center", minHeight: 18 }}>{children}</div>; }

function PolicyDocModal({ form, uw, coverages, premium, quoteRef, selectedName, onClose }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(14,27,61,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500, padding: 24 }}>
      <div style={{ background: "#fff", width: 640, maxHeight: "88vh", overflowY: "auto", borderRadius: 10, padding: "40px 48px" }} id="policy-doc">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: `2px solid ${T.navy}`, paddingBottom: 16, marginBottom: 24 }}>
          <div>
            <div style={{ fontSize: 22, fontWeight: 800, color: T.navy }}>SME PRODUCT QUOTATION</div>
            <div style={{ fontSize: 11, color: T.muted, letterSpacing: "0.05em" }}>SME PACKAGE POLICY SCHEDULE</div>
          </div>
          <div style={{ textAlign: "right", fontSize: 11.5, color: T.muted, fontWeight: 600 }}>{quoteRef}<br />{form.Effective_Date}</div>
        </div>
        <DocRow label="Insured Business Activity" value={form.Business_Activity} />
        <DocRow label="Industry Segment" value={form.Industry_Segment} />
        <DocRow label="Distribution Channel" value={`${form.Distribution_Channel} (${form.Agent_Tier} tier)`} />
        <DocRow label="Policy Term" value={`${form.Policy_Term_Months} months, effective ${form.Effective_Date}`} />
        <DocRow label="Package" value={selectedName} />
        <DocRow label="Underwriting Decision" value={uw?.UW_Status} />
        <div style={{ marginTop: 22, marginBottom: 10, fontSize: 12, fontWeight: 700, color: T.navy, textTransform: "uppercase", letterSpacing: "0.05em" }}>Coverage Schedule</div>
        <CoverageTable rows={coverages} />
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 18, paddingTop: 16, borderTop: `1px solid ${T.mid}` }}>
          <div style={{ textAlign: "right" }}>
            <div style={{ fontSize: 11, color: T.muted }}>TOTAL ANNUAL PREMIUM</div>
            <div style={{ fontSize: 20, fontWeight: 800, color: T.navy }}>{fmtMoney(premium)}</div>
          </div>
        </div>
        <div style={{ fontSize: 10.5, color: T.muted, marginTop: 28, lineHeight: 1.6 }}>
          This is a demonstration document generated for illustrative purposes as part of the Coherent Spark SME digital platform proof of concept. It is not a binding policy document.
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 26 }}>
          <GhostButton onClick={onClose}>Close</GhostButton>
          <PrimaryButton onClick={() => window.print()}>Print / Save as PDF</PrimaryButton>
        </div>
      </div>
    </div>
  );
}
function DocRow({ label, value }) {
  return <div style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${T.mid}`, fontSize: 13 }}><span style={{ color: T.muted }}>{label}</span><span style={{ fontWeight: 700, color: T.text }}>{value}</span></div>;
}

// ─── AI Assistant sidebar ────────────────────────────────────────────────────
// A collapsible chat drawer that can (1) fill in form fields from natural-language
// descriptions of the business, (2) ask clarifying questions instead of the agent
// hunting through dropdowns, and (3) explain package/coverage differences. Calls
// Claude directly from the browser (per the artifact "AI-powered app" pattern) and
// asks for a strict {reply, updates} JSON shape so field updates can be applied
// deterministically rather than parsed out of free text.
function buildAssistantSystemPrompt(form, step, offers, activityOptions) {
  return `You are a helper assistant embedded in a Coherent SME insurance quoting tool. The user is an insurance agent filling out a form to get a business quote.

Your jobs, in priority order:
1. If the user describes their business or risk details in natural language (e.g. "supermarket, $2M turnover, 15 staff, no sprinklers"), extract whatever field values you can confidently identify and put them in "updates".
2. If something important is missing or ambiguous to proceed, ask ONE short clarifying question in "reply" rather than guessing.
3. If the user asks a question (e.g. "what's the difference between the packages?", "what does BI cover?"), answer helpfully and concisely in "reply" using the context below, and leave "updates" empty.
4. Keep "reply" conversational and skimmable. You can use **bold** for key terms/numbers, _underline_ for emphasis, and "- " bullet lines for lists (e.g. comparing packages or listing reasons) — these render properly in the chat. Use them when they genuinely aid readability (e.g. comparing 2-3 things), not for a one-line answer.

Valid field names (use these exact keys) and allowed values:
- Distribution_Channel: Broker | Agent | Direct
- Agent_Tier: Standard | Silver | Gold
- Effective_Date: YYYY-MM-DD
- Policy_Term_Months: 12 | 24 | 36
- Industry_Segment: ${INDUSTRY_SEGMENTS.join(" | ")}
- Business_Activity: MUST be one of the activities currently valid for the SELECTED Industry_Segment (see below) — never set an activity that belongs to a different segment. If the user names an activity that belongs to a different segment than what's currently selected, set Industry_Segment to the matching one first (in the same "updates"), and only set Business_Activity if you're confident it's valid there; otherwise ask instead of guessing.
  Activities valid for the CURRENT Industry_Segment ("${form.Industry_Segment}"): ${activityOptions && activityOptions.length ? activityOptions.join(", ") : "(none loaded yet for this segment)"}
- Years_In_Business, Employee_Count, Premises_Count: whole numbers
- Annual_Turnover, Property_Sum_Insured: numbers in USD, no commas or symbols
- Building_Type: Detached Building | Office / Strata Unit | Mall / Retail Unit | Industrial Unit
- Construction: Concrete / Brick | Mixed | Light / Metal
- Fire_Protection: Full Sprinklers | Alarm & Extinguishers | Basic Extinguishers | None
- Flood_Zone: Low | Moderate | High
- Cyber_Exposure: Low | Medium | High
- Prior_Claims_3Y: None | 1 | 2 | 3+
- Hot_Works, Cooking_Deep_Frying, Stock_In_Transit, Manual_Work_Away: Yes | No
- Commission_Pct, Manual_Adjustment_Pct: decimals, e.g. 0.12 means 12%

Current form values: ${JSON.stringify(form)}
Current step: ${step}
${offers && offers.length ? `Current package options (for explaining differences): ${JSON.stringify(offers)}` : ""}

Respond with ONLY a raw JSON object (no markdown code fences, no preamble), in exactly this shape:
{"reply": "your message to the user", "updates": {"Field_Name": value, ...}}
Only include a field in "updates" if you are confident about it; omit anything uncertain. Never claim in "reply" that something was updated unless you actually included it in "updates".`;
}
function sanitizeAssistantUpdates(updates, form, activityOptions, industryMapping) {
  const allowed = {
    Distribution_Channel: OPTS.Distribution_Channel, Agent_Tier: OPTS.Agent_Tier,
    Building_Type: OPTS.Building_Type, Construction: OPTS.Construction, Fire_Protection: OPTS.Fire_Protection,
    Flood_Zone: OPTS.Flood_Zone, Cyber_Exposure: OPTS.Cyber_Exposure, Prior_Claims_3Y: OPTS.Prior_Claims_3Y,
    Hot_Works: ["Yes", "No"], Cooking_Deep_Frying: ["Yes", "No"], Stock_In_Transit: ["Yes", "No"], Manual_Work_Away: ["Yes", "No"],
    Industry_Segment: INDUSTRY_SEGMENTS,
    // If the full mapping loaded, any known activity is acceptable (segment gets
    // auto-corrected to match afterward); otherwise fall back to the current segment's list.
    Business_Activity: industryMapping?.byActivity ? Object.keys(industryMapping.byActivity) : activityOptions,
  };
  const numeric = ["Years_In_Business", "Employee_Count", "Premises_Count", "Annual_Turnover", "Property_Sum_Insured", "Commission_Pct", "Manual_Adjustment_Pct", "Policy_Term_Months"];
  const applied = {};
  const rejected = [];
  Object.entries(updates || {}).forEach(([k, v]) => {
    if (!(k in form)) { rejected.push({ field: k, reason: "not a recognized field" }); return; }
    if (allowed[k] && !allowed[k].includes(v)) {
      rejected.push({ field: k, reason: k === "Business_Activity" ? "not a valid activity for the current Industry Segment" : "not a valid value" });
      return;
    }
    if (numeric.includes(k)) {
      const n = Number(v);
      if (isNaN(n)) { rejected.push({ field: k, reason: "not a valid number" }); return; }
      applied[k] = n; return;
    }
    applied[k] = v;
  });
  // Auto-correct Industry_Segment to match a resolved Business_Activity, using the
  // authoritative mapping rather than trusting the model to pair them correctly.
  if (applied.Business_Activity && industryMapping?.byActivity?.[applied.Business_Activity]) {
    applied.Industry_Segment = industryMapping.byActivity[applied.Business_Activity];
  }
  return { applied, rejected };
}
// Lightweight, dependency-free markdown-lite renderer for chat bubbles: **bold**,
// _underline_, and "- " bullet lists. Builds React elements directly (no
// dangerouslySetInnerHTML) so there's no HTML-injection risk from model output.
function renderInline(str, keyPrefix) {
  const parts = [];
  let remaining = str;
  const regex = /(\*\*(.+?)\*\*|_(.+?)_)/;
  let key = 0;
  while (remaining.length) {
    const m = remaining.match(regex);
    if (!m) { parts.push(remaining); break; }
    if (m.index > 0) parts.push(remaining.slice(0, m.index));
    if (m[2] !== undefined) parts.push(<strong key={`${keyPrefix}-${key++}`}>{m[2]}</strong>);
    else if (m[3] !== undefined) parts.push(<u key={`${keyPrefix}-${key++}`}>{m[3]}</u>);
    remaining = remaining.slice(m.index + m[0].length);
  }
  return parts;
}
function renderRichText(text) {
  if (!text) return null;
  const lines = String(text).split("\n");
  const elements = [];
  let listBuffer = [];
  const flushList = () => {
    if (listBuffer.length) {
      elements.push(
        <ul key={`ul-${elements.length}`} style={{ margin: "4px 0", paddingLeft: 18 }}>
          {listBuffer.map((li, i) => <li key={i} style={{ marginBottom: 2 }}>{renderInline(li, `li-${elements.length}-${i}`)}</li>)}
        </ul>
      );
      listBuffer = [];
    }
  };
  lines.forEach((line, i) => {
    const trimmed = line.trim();
    if (/^[-•]\s+/.test(trimmed)) {
      listBuffer.push(trimmed.replace(/^[-•]\s+/, ""));
    } else {
      flushList();
      if (trimmed === "") elements.push(<div key={`sp-${i}`} style={{ height: 4 }} />);
      else elements.push(<div key={`ln-${i}`}>{renderInline(trimmed, `ln-${i}`)}</div>);
    }
  });
  flushList();
  return elements;
}
async function callAssistant(system, historyMessages) {
  try {
    const res = await fetch("/api/assistant", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ system, messages: historyMessages }),
    });
    const data = await res.json();
    const raw = data?.choices?.[0]?.message?.content || "";
    const cleaned = raw.replace(/```json|```/g, "").trim();
    try { return JSON.parse(cleaned); } catch { return { reply: raw || "Sorry, I didn't quite catch that.", updates: {} }; }
  } catch {
    return { reply: "I couldn't reach the assistant just now — please try again.", updates: {} };
  }
}
function AIAssistant({ form, setForm, step, offers, activityOptions, industryMapping }) {
  const [open, setOpen] = useState(false);
  const initialGreeting = { role: "assistant", text: "Hi! Tell me about the business in your own words and I'll fill in the form — e.g. \"Supermarket, $2M turnover, 15 staff, no sprinklers.\" You can also ask me things like \"what's the difference between the packages?\"" };
  const [messages, setMessages] = useState([initialGreeting]);
  const [input, setInput] = useState("");
  const [loading, setLoading] = useState(false);
  const scrollRef = useRef(null);

  useEffect(() => { if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight; }, [messages, open, loading]);

  function newConversation() {
    setMessages([initialGreeting]);
    setInput("");
  }

  async function send() {
    const text = input.trim();
    if (!text || loading) return;
    const next = [...messages, { role: "user", text }];
    setMessages(next);
    setInput("");
    setLoading(true);
    const system = buildAssistantSystemPrompt(form, step, offers, activityOptions);
    // (industryMapping intentionally not sent to the model — resolution happens deterministically in sanitizeAssistantUpdates instead)
    const history = next.slice(-8).map(m => ({ role: m.role === "user" ? "user" : "assistant", content: m.text }));
    const parsed = await callAssistant(system, history);
    const { applied, rejected } = sanitizeAssistantUpdates(parsed.updates, form, activityOptions, industryMapping);
    if (Object.keys(applied).length) setForm(f => ({ ...f, ...applied }));
    let replyText = parsed.reply || "Done.";
    if (rejected.length) {
      replyText += `\n\n⚠ Couldn't apply: ${rejected.map(r => `${humanize(r.field)} (${r.reason})`).join(", ")}.`;
    }
    setMessages(m => [...m, { role: "assistant", text: replyText, updates: applied }]);
    setLoading(false);
  }

  return (
    <>
      <button onClick={() => setOpen(o => !o)} style={{
        position: "fixed", right: open ? 340 : 0, top: "50%", transform: "translateY(-50%)",
        background: T.navy, color: "#fff", border: "none", borderRadius: "8px 0 0 8px", padding: "18px 9px",
        cursor: "pointer", zIndex: 200, writingMode: "vertical-rl", fontSize: 11.5, fontWeight: 700,
        letterSpacing: "0.06em", transition: "right 0.25s ease", boxShadow: "-2px 0 10px rgba(14,27,61,0.15)",
      }}>{open ? "CLOSE" : "AI ASSISTANT"}</button>
      <div style={{
        position: "fixed", top: 0, right: 0, height: "100vh", width: 340, background: T.white,
        boxShadow: "-6px 0 24px rgba(14,27,61,0.14)", transform: `translateX(${open ? 0 : 340}px)`,
        transition: "transform 0.25s ease", zIndex: 190, display: "flex", flexDirection: "column",
      }}>
        <div style={{ background: T.navy, color: "#fff", padding: "16px 18px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>AI Assistant</div>
            <div style={{ fontSize: 10.5, color: "rgba(255,255,255,0.55)", marginTop: 2 }}>Fills in fields, asks questions, explains coverage</div>
          </div>
          <button onClick={newConversation} title="Start a new conversation" style={{
            background: "rgba(255,255,255,0.12)", border: "none", borderRadius: 6, color: "#fff",
            fontSize: 10.5, fontWeight: 700, padding: "6px 10px", cursor: "pointer", flexShrink: 0, marginLeft: 8,
          }}>New Chat</button>
        </div>
        <div ref={scrollRef} style={{ flex: 1, overflowY: "auto", padding: 16, display: "flex", flexDirection: "column", gap: 10 }}>
          {messages.map((m, i) => (
            <div key={i} style={{
              alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%",
              background: m.role === "user" ? T.navy : T.mid, color: m.role === "user" ? "#fff" : T.text,
              padding: "9px 12px", borderRadius: 10, fontSize: 13, lineHeight: 1.45,
              textAlign: m.role === "user" ? "right" : "left",
            }}>
              {renderRichText(m.text)}
              {m.updates && Object.keys(m.updates).length > 0 && (
                <div style={{ marginTop: 6, fontSize: 10.5, opacity: 0.7, borderTop: "1px solid rgba(0,0,0,0.1)", paddingTop: 5 }}>
                  Updated: {Object.keys(m.updates).map(humanize).join(", ")}
                </div>
              )}
            </div>
          ))}
          {loading && <div style={{ alignSelf: "flex-start", fontSize: 12, color: T.muted, display: "flex", alignItems: "center", gap: 6 }}><Spinner /> Thinking…</div>}
        </div>
        <div style={{ padding: 12, borderTop: `1px solid ${T.mid}`, display: "flex", gap: 8 }}>
          <input value={input} onChange={e => setInput(e.target.value)} onKeyDown={e => e.key === "Enter" && send()}
            placeholder="Tell me about the business…" disabled={loading}
            style={{ ...inputBase(false), flex: 1, height: 38 }} />
          <PrimaryButton onClick={send} disabled={loading || !input.trim()}>Send</PrimaryButton>
        </div>
      </div>
    </>
  );
}

// ─── API Log (password-protected) ───────────────────────────────────────────
const API_LOG_PASSWORD = "Coherent!";

function ApiLogTab({ onClick }) {
  return (
    <button onClick={onClick} style={{
      position: "fixed", left: 0, top: "50%", transform: "translateY(-50%)",
      background: T.navy, color: "#fff", border: "none", borderRadius: "0 8px 8px 0", padding: "18px 9px",
      cursor: "pointer", zIndex: 200, writingMode: "vertical-rl", fontSize: 11.5, fontWeight: 700,
      letterSpacing: "0.06em", boxShadow: "2px 0 10px rgba(14,27,61,0.15)",
    }}>API LOG</button>
  );
}

function ApiLogModal({ log, onClose }) {
  const [unlocked, setUnlocked] = useState(false);
  const [pwd, setPwd] = useState("");
  const [pwdError, setPwdError] = useState(false);
  const [expandedId, setExpandedId] = useState(null);

  function tryUnlock() {
    if (pwd === API_LOG_PASSWORD) { setUnlocked(true); setPwdError(false); }
    else { setPwdError(true); }
  }

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(14,27,61,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 500, padding: 24 }}>
      <div style={{ background: "#fff", width: unlocked ? 900 : 380, maxWidth: "94vw", maxHeight: "86vh", overflowY: "auto", borderRadius: 10, padding: unlocked ? "26px 30px" : "36px 40px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <div style={{ fontFamily: FONT, fontSize: 18, fontWeight: 800, color: T.navy }}>API Call Log</div>
            <div style={{ fontSize: 11.5, color: T.muted, marginTop: 2 }}>
              {unlocked ? `${log.length} call${log.length === 1 ? "" : "s"} recorded this session` : "Password protected"}
            </div>
          </div>
          <GhostButton onClick={onClose}>Close</GhostButton>
        </div>

        {!unlocked ? (
          <div>
            <Field label="Password" error={pwdError ? "Incorrect password" : undefined}>
              <input
                type="password"
                value={pwd}
                autoFocus
                onChange={e => { setPwd(e.target.value); setPwdError(false); }}
                onKeyDown={e => e.key === "Enter" && tryUnlock()}
                style={inputBase(pwdError)}
              />
            </Field>
            <PrimaryButton onClick={tryUnlock} full>Unlock</PrimaryButton>
          </div>
        ) : (
          <>
            <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginBottom: 14 }}>
              <GhostButton onClick={() => downloadApiLogCSV(log)} disabled={!log.length}>Download CSV</GhostButton>
            </div>
            {!log.length ? (
              <div style={{ textAlign: "center", padding: "40px 0", color: T.muted, fontSize: 13 }}>No API calls recorded yet this session.</div>
            ) : (
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                  <tr>
                    {["Time", "Purpose", "Category", "Status", "Duration", ""].map(h => (
                      <th key={h} style={{ padding: "8px 10px", background: T.navy, color: "#fff", fontSize: 10.5, fontWeight: 700, letterSpacing: "0.05em", textTransform: "uppercase", textAlign: "left" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[...log].reverse().map(e => (
                    <Fragment key={e.id}>
                      <tr>
                        <td style={{ padding: "8px 10px", fontSize: 12, borderBottom: `1px solid ${T.mid}`, whiteSpace: "nowrap" }}>{new Date(e.timestamp).toLocaleTimeString()}</td>
                        <td style={{ padding: "8px 10px", fontSize: 12, borderBottom: `1px solid ${T.mid}`, fontWeight: 600 }}>{e.purpose}</td>
                        <td style={{ padding: "8px 10px", fontSize: 11.5, borderBottom: `1px solid ${T.mid}`, color: T.muted }}>{e.serviceCategory}</td>
                        <td style={{ padding: "8px 10px", fontSize: 12, borderBottom: `1px solid ${T.mid}` }}><Pill tone={e.status === "success" ? "green" : "red"}>{e.status}</Pill></td>
                        <td style={{ padding: "8px 10px", fontSize: 12, borderBottom: `1px solid ${T.mid}` }}>{e.durationMs}ms</td>
                        <td style={{ padding: "8px 10px", fontSize: 11.5, borderBottom: `1px solid ${T.mid}` }}>
                          <button onClick={() => setExpandedId(id => id === e.id ? null : e.id)} style={{ background: "none", border: "none", color: T.navy, fontWeight: 700, cursor: "pointer", fontSize: 11.5, fontFamily: FONT }}>
                            {expandedId === e.id ? "Hide" : "View"}
                          </button>
                        </td>
                      </tr>
                      {expandedId === e.id && (
                        <tr>
                          <td colSpan={6} style={{ padding: "12px 14px", background: T.light, borderBottom: `1px solid ${T.mid}` }}>
                            <div style={{ fontSize: 10.5, fontWeight: 700, color: T.muted, marginBottom: 4, textTransform: "uppercase" }}>Request</div>
                            <pre style={{ fontSize: 11, whiteSpace: "pre-wrap", wordBreak: "break-word", background: T.white, padding: 10, borderRadius: 6, border: `1px solid ${T.mid}`, marginBottom: 10 }}>{e.requestBody}</pre>
                            <div style={{ fontSize: 10.5, fontWeight: 700, color: T.muted, marginBottom: 4, textTransform: "uppercase" }}>Response</div>
                            <pre style={{ fontSize: 11, whiteSpace: "pre-wrap", wordBreak: "break-word", background: T.white, padding: 10, borderRadius: 6, border: `1px solid ${T.mid}` }}>{e.responseBody}</pre>
                          </td>
                        </tr>
                      )}
                    </Fragment>
                  ))}
                </tbody>
              </table>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// ─── Root app ────────────────────────────────────────────────────────────────
export default function App() {
  const [step, setStep] = useState("business");
  const [form, setForm] = useState(DEFAULT_FORM);
  const [live, setLive] = useState(false);
  const [checkingConn, setCheckingConn] = useState(true);
  const [schema, setSchema] = useState(null);
  const [serviceMeta, setServiceMeta] = useState(null);
  const [errors, setErrors] = useState({});
  const [shakeKey, setShakeKey] = useState(0);

  const [activityOptions, setActivityOptions] = useState(BUSINESS_ACTIVITIES);
  const [industryMapping, setIndustryMapping] = useState(STATIC_INDUSTRY_MAPPING);
  const [activityLoading, setActivityLoading] = useState(false);
  const [resolvedActivity, setResolvedActivity] = useState({ code: null, appetite: null, trigger: null, ratingClass: null, occupancyGroup: null });
  const [resolvedLoading, setResolvedLoading] = useState(false);
  const segmentMounted = useRef(false);

  const [uwLoading, setUwLoading] = useState(false);
  const [overlay, setOverlay] = useState(false);
  const [uw, setUw] = useState(null);
  const [offers, setOffers] = useState([]);
  const [selectedPkg, setSelectedPkg] = useState(null);
  const [mode, setMode] = useState("quick");
  const [reachable, setReachable] = useState(["business"]);

  const [endorsement, setEndorsement] = useState(true);
  useEffect(() => {
    const end1 = selectedPkg?.groups?.find(g => g.name === "Endorsement 1")?.include;
    if (end1 === "None") setEndorsement(false);
    else if (end1 === "Required") setEndorsement(true);
  }, [selectedPkg]);
  const [bindResult, setBindResult] = useState(null);
  const [bindIssues, setBindIssues] = useState(null);
  const [bindSource, setBindSource] = useState("quick");
  const [bindLoading, setBindLoading] = useState(false);

  const [flexSelections, setFlexSelections] = useState({
    Property: true, Liability: true, "Employers Liability": true,
    "Business Interruption": false, Cyber: false, "Endorsement 1": false,
  });
  const [flexPriced, setFlexPriced] = useState(null);
  const [flexLoading, setFlexLoading] = useState(false);
  const [flexRequirements, setFlexRequirements] = useState(null);
  const [flexReqLoading, setFlexReqLoading] = useState(false);
  const [flexPackageCode, setFlexPackageCode] = useState(null);

  const [ticker, setTicker] = useState({ source: "loading", status: null, indicativeAction: "", reasons: "", low: null, high: null });
  const tickerTimer = useRef(null);

  const [apiLog, setApiLog] = useState([]);
  const [apiLogOpen, setApiLogOpen] = useState(false);
  const logCall = (entry) => setApiLog(l => [...l, entry].slice(-300));

  const quoteRef = useMemo(() => `SME-${Math.abs(hashCode(form.Business_Activity + form.Effective_Date)).toString().slice(0, 6)}`, [form.Business_Activity, form.Effective_Date]);

  useEffect(() => {
    (async () => {
      const s = await fetchValidationSchema({}, {}, logCall);
      setSchema(s);
      setServiceMeta(s?.__meta || null);
      setLive(!!s);
      setCheckingConn(false);
    })();
  }, []);

  // Fetch the full Industry Segment <-> Business Activity mapping once, so segment
  // changes can look up the activity list instantly instead of a network round-trip
  // each time, and the AI Assistant can resolve the right segment when only given an
  // activity name.
  useEffect(() => {
    (async () => {
      const res = await callSpark("#Default#", { "Flex_Quote_Bind.Coverages": [] }, "Industry Mapping", { onLog: logCall });
      if (res.ok) {
        const parsed = parseIndustryMapping(res.data["Industry_and_Business_Mapping"]);
        if (parsed) setIndustryMapping(parsed);
      }
    })();
  }, []);

  // Prefer the cached full Industry_and_Business_Mapping (instant, no network call) if
  // it loaded successfully; otherwise fall back to the confirmed live per-segment
  // validation call (Dynamic + requested_output), and finally the captured static table.
  const lastSegment = useRef(form.Industry_Segment);
  useEffect(() => {
    const previousActivity = form.Business_Activity;
    let cancelled = false;
    const segmentChanged = segmentMounted.current && lastSegment.current !== form.Industry_Segment;
    lastSegment.current = form.Industry_Segment;

    const cached = industryMapping?.bySegment?.[form.Industry_Segment];
    if (cached && cached.length) {
      setActivityOptions(cached);
      setActivityLoading(false);
    } else {
      setActivityLoading(true);
      (async () => {
        const s = await fetchValidationSchema(
          { Industry_Segment: form.Industry_Segment, Business_Activity: previousActivity || "" },
          { validation_type: "Dynamic", requested_output: "Business_Activity" },
          logCall
        );
        if (cancelled) return;
        const liveOpts = s?.Business_Activity?.options;
        setActivityOptions(liveOpts && liveOpts.length ? liveOpts : (ACTIVITIES_BY_SEGMENT[form.Industry_Segment] || []));
        setActivityLoading(false);
      })();
    }

    if (segmentChanged) {
      setForm(f => ({ ...f, Business_Activity: "" }));
      setResolvedActivity({ code: null, appetite: null, trigger: null, ratingClass: null, occupancyGroup: null });
    }
    segmentMounted.current = true;
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.Industry_Segment, industryMapping]);

  // Rating_Class / Activity_Code / Occupancy_Group / Activity_Appetite / Activity_UW_Trigger
  // are plain computed formula cells in the workbook, not Excel dropdown-validated cells —
  // so the /validation "Dynamic" mechanism (which is specifically about resolving dropdown
  // options/defaults) never actually applies to them, regardless of what's in the workbook.
  // The correct approach: a real execute call with these 5 fields left blank as inputs,
  // reading back whatever the workbook computes for them as outputs.
  useEffect(() => {
    if (!form.Business_Activity) { setResolvedActivity({ code: null, appetite: null, trigger: null, ratingClass: null, occupancyGroup: null }); return; }
    let cancelled = false;
    setResolvedLoading(true);
    (async () => {
      const inputs = { Industry_Segment: form.Industry_Segment, Business_Activity: form.Business_Activity };
      const res = await callSpark("#Default#", inputs, "Activity Resolution", { onLog: logCall });
      if (cancelled) return;
      const code = res.ok ? (res.data?.Activity_Code ?? null) : null;
      const appetite = res.ok ? (res.data?.Activity_Appetite ?? null) : null;
      const trigger = res.ok ? (res.data?.Activity_UW_Trigger ?? null) : null;
      const ratingClass = res.ok ? (res.data?.Rating_Class ?? null) : null;
      const occupancyGroup = res.ok ? (res.data?.Occupancy_Group ?? null) : null;
      if (code || ratingClass || occupancyGroup) {
        setResolvedActivity({ code, appetite, trigger, ratingClass, occupancyGroup });
      } else {
        const fb = ACTIVITY_TABLE[form.Business_Activity];
        setResolvedActivity(fb
          ? { code: fb.code, appetite: fb.appetite, trigger: fb.trigger, ratingClass: fb.ratingClass, occupancyGroup: fb.occupancyGroup }
          : { code: null, appetite: null, trigger: null, ratingClass: null, occupancyGroup: null });
      }
      setResolvedLoading(false);
    })();
    return () => { cancelled = true; };
  }, [form.Business_Activity, form.Industry_Segment]);

  const resolvedInputs = () => ({
    Activity_Code: resolvedActivity.code, Activity_Appetite: resolvedActivity.appetite, Activity_UW_Trigger: resolvedActivity.trigger,
    Rating_Class: resolvedActivity.ratingClass, Occupancy_Group: resolvedActivity.occupancyGroup,
  });
  // Real Spark calls collapse into a cascade of #CALC!/#N/A errors if these resolved
  // fields are null — so every live call below is guarded on this rather than silently
  // sending nulls and getting back a broken response.
  const activityResolved = !!(resolvedActivity.code && resolvedActivity.ratingClass && resolvedActivity.occupancyGroup);

  const goto = (k) => { if (reachable.includes(k)) setStep(k); };
  const markReachable = (k) => setReachable(r => r.includes(k) ? r : [...r, k]);

  function validateStep(stepKey) {
    const req = REQUIRED_BY_STEP[stepKey] || [];
    const next = {};
    req.forEach(f => { if (form[f] === "" || form[f] === null || form[f] === undefined) next[f] = "This field is required."; });
    setErrors(next);
    if (Object.keys(next).length) { setShakeKey(k => k + 1); return false; }
    return true;
  }

  // BUG FIX: the captured API doc sample for this v2 service shows that when
  // Flex_Quote_Bind_Coverages is omitted, Spark returns a validation error
  // ("Expected array" on inputs.Flex_Quote_Bind_Coverages) AND every
  // Quick_Quote_Bind coverage row comes back zeroed/blank — exactly the "Bind
  // Quote shows $0 for everything" symptom. Even though the field is marked
  // optional in the schema, the real workbook behavior treats a missing value
  // as a type error that appears to break other computed outputs in the same
  // call. Sending an empty array by default (real Flex selections override it
  // when actually binding Flex) avoids that error entirely.
  function baseInputs(extra = {}) {
    return { ...form, ...resolvedInputs(), "Flex_Quote_Bind.Coverages": [], ...extra };
  }

  async function runAppetiteAndQuickQuote() {
    if (!validateStep("risk")) return;
    setUwLoading(true); setOverlay(true);
    let res = { ok: false };
    if (activityResolved) {
      [res] = await Promise.all([callSpark("#Default#,Quick_Quote", baseInputs(), "Appetite Check", { onLog: logCall }), new Promise(r => setTimeout(r, 1700))]);
    } else {
      await new Promise(r => setTimeout(r, 1200)); // keep the reasoning overlay feeling deliberate
    }
    let outputs;
    if (res.ok) { outputs = res.data; setLive(true); } else { outputs = { ...simulateUnderwriting(form), ...simulateQuickQuote(form) }; setLive(false); }
    setUw({ UW_Status: outputs.UW_Status, UW_Reasons: outputs.UW_Reasons, UW_Indicative_Actions: outputs.UW_Indicative_Actions, UW_Commercial_Action: outputs.UW_Commercial_Action });
    const parsedOffers = pivotOffers(outputs["Quick_Quote.Offers"]);
    setOffers(parsedOffers);
    setSelectedPkg(parsedOffers[Math.floor(parsedOffers.length / 2)] || null);
    setUwLoading(false); setOverlay(false);
    setStep("appetite"); markReachable("appetite");
  }

  async function runQuickBind(purpose = "Bind") {
    setBindLoading(true);
    let res = { ok: false };
    if (activityResolved) {
      const inputs = baseInputs({ "Quick_Quote_Bind.Package_Code": selectedPkg?.code, "Quick_Quote_Bind.END1": endorsement });
      res = await callSpark("#Default#,Quick_Quote_Bind", inputs, purpose, { onLog: logCall });
    }
    const outputs = res.ok ? res.data : simulateQuickBind(form, selectedPkg?.code);
    setLive(res.ok);
    setBindIssues(res.ok ? { warnings: res.warnings, errors: res.errors } : null);
    setBindResult(outputs); setBindSource("quick"); setBindLoading(false);
  }

  // Auto-bind as soon as the standard-package Bind step is reached — previously this
  // required a separate manual "Bind Package" click that was easy to miss, leaving the
  // page looking stuck with a disabled Continue button and no visible next action.
  useEffect(() => {
    if (step === "bind" && bindSource === "quick" && !bindResult && selectedPkg) runQuickBind();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [step, bindSource]);

  async function loadFlexRequirements() {
    if (!activityResolved) { setFlexReqLoading(false); return; }
    setFlexReqLoading(true);
    const res = await callSpark("#Default#,Flex_Quote", baseInputs(), "Flex Configuration", { onLog: logCall });
    if (res.ok) {
      const list = res.data["Flex_Quote.List"];
      const reqs = parseFlexRequirements(list);
      if (reqs) setFlexRequirements(reqs);
      const code = parseFlexPackageCode(list);
      if (code) setFlexPackageCode(code);
      setLive(true);
    }
    setFlexReqLoading(false);
  }

  async function runFlexBind() {
    setFlexLoading(true);
    const coveragesInput = [
      { Key: "Package Code", Value: flexPackageCode || FLEX_PACKAGE_CODE_FALLBACK, Include: true },
      ...Object.entries(flexSelections).map(([cov, include]) => ({ Key: cov, Value: FLEX_FIXED_CODE[cov], Include: include || FLEX_MANDATORY.includes(cov) })),
    ];
    let res = { ok: false };
    if (activityResolved) {
      const inputs = baseInputs({ "Flex_Quote_Bind.Coverages": coveragesInput });
      res = await callSpark("#Default#,Flex_Quote_Bind", inputs, "Reprice Flex Quote", { onLog: logCall });
    }
    const outputs = res.ok ? res.data : simulateFlexBind(form, flexSelections);
    setLive(res.ok);
    setBindIssues(res.ok ? { warnings: res.warnings, errors: res.errors } : null);
    setFlexPriced(outputs); setFlexLoading(false);
  }

  function resetAll() {
    setStep("business"); setForm(DEFAULT_FORM); setUw(null); setOffers([]); setSelectedPkg(null); setErrors({});
    setMode("quick"); setReachable(["business"]); setBindResult(null); setBindIssues(null); setFlexPriced(null); setFlexRequirements(null);
    setFlexSelections({
      Property: true, Liability: true, "Employers Liability": true,
      "Business Interruption": false, Cyber: false, "Endorsement 1": false,
    });
  }

  const selectedName = bindSource === "flex" ? "Flex-Built Package" : selectedPkg?.name;
  const showTicker = ["business", "risk"].includes(step);

  async function runTickerCheck(purpose = "Quotation") {
    if (!activityResolved) {
      setTicker({ source: "unresolved", status: null, indicativeAction: "", reasons: "", low: null, high: null });
      return;
    }
    setTicker(t => ({ ...t, source: "loading" }));
    const res = await callSpark("#Default#,Quick_Quote", baseInputs(), purpose, { onLog: logCall });
    if (res.ok) {
      const offers = pivotOffers(res.data["Quick_Quote.Offers"]);
      const prices = offers.map(o => Number(o.premium)).filter(n => !isNaN(n));
      setTicker({
        source: "live", status: res.data.UW_Status, indicativeAction: res.data.UW_Indicative_Actions, reasons: res.data.UW_Reasons,
        low: prices.length ? Math.min(...prices) : null, high: prices.length ? Math.max(...prices) : null,
      });
      setLive(true);
    } else {
      const { base, riskLoad } = ratingBaseAndLoad(form);
      const mid = Math.max(650, Math.round(base * riskLoad));
      const sim = simulateUnderwriting(form);
      setTicker({ source: "estimate", status: sim.UW_Status, indicativeAction: sim.UW_Indicative_Actions, reasons: sim.UW_Reasons, low: Math.round(mid * 0.68), high: Math.round(mid * 1.45) });
      setLive(false);
    }
  }

  useEffect(() => {
    if (!showTicker) return;
    setTicker(t => ({ ...t, source: "loading" }));
    if (tickerTimer.current) clearTimeout(tickerTimer.current);
    tickerTimer.current = setTimeout(runTickerCheck, 500);
    return () => clearTimeout(tickerTimer.current);
    // Deliberately keyed on Business_Activity / Industry_Segment only — these are the two
    // fields that meaningfully move the preliminary appetite/pricing signal early in the
    // form, so we call Spark only when either changes rather than on every keystroke
    // across the whole Business/Risk profile (keeps the demo from being "too chatty").
    // A manual "Check Eligibility Now" button lets the agent force a fresh, fully-accurate
    // check against every current field whenever they want, without waiting on this.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [form.Business_Activity, form.Industry_Segment, resolvedActivity.code, showTicker]);

  return (
    <div style={{ minHeight: "100vh", background: T.light, fontFamily: FONT, color: T.text }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes shimmer { 0% { background-position: 100% 0; } 100% { background-position: 0 0; } }
        @keyframes fadeSlide { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes shake { 10%,90%{transform:translateX(-1px);} 20%,80%{transform:translateX(2px);} 30%,50%,70%{transform:translateX(-4px);} 40%,60%{transform:translateX(4px);} }
        select:focus, input:focus { outline: none; }
        @media print { #policy-doc { box-shadow: none !important; } }
      `}</style>

      <div style={{ background: T.white, padding: "0 32px", height: 76, display: "grid", gridTemplateColumns: "1fr auto 1fr", alignItems: "center", borderBottom: `1px solid ${T.mid}`, boxShadow: "0 1px 8px rgba(14,27,61,0.05)", position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16, justifySelf: "start" }}>
          <LogoBox src="/coherent-logo.svg" fallback="COHERENT" />
        </div>
        <div style={{ textAlign: "center", justifySelf: "center" }}>
          <div style={{ fontFamily: FONT, fontSize: 23, fontWeight: 800, color: T.navy, letterSpacing: "0.01em", lineHeight: 1.15 }}>SME Product Quote</div>
          <div style={{ fontSize: 10.5, color: T.muted, letterSpacing: "0.09em", textTransform: "uppercase", marginTop: 3, fontWeight: 600 }}>Coherent Spark Demo</div>
        </div>
        <div style={{ justifySelf: "end" }}>
          <ConnectionBadge live={live} checking={checkingConn} />
        </div>
      </div>

      <div style={{ maxWidth: 1180, margin: "0 auto", padding: "28px 24px 60px", display: "flex", gap: 32, alignItems: "flex-start" }}>
        <div style={{ flex: 1, minWidth: 0, maxWidth: 820, textAlign: "left" }}>
          <div style={{ marginBottom: 22 }}><StepBar step={step} reachable={reachable} onJump={goto} /></div>
          {!live && !checkingConn && (
            <Banner tone="error" title="✕ Spark API unreachable — using built-in demo calculations.">
              Figures are illustrative until connectivity is restored. If this persists while Postman/curl calls succeed, it's worth checking the browser console for the actual fetch error (CORS and plain network failures both show as blocked responses in Network tab, but only the console reveals which).
            </Banner>
          )}
          <div key={shakeKey} style={Object.keys(errors).length ? { animation: "shake 0.4s" } : undefined}>
            <FadeIn k={step}>
              {step === "business" && <StepBusiness form={form} setForm={setForm} schema={schema} errors={errors} activityOptions={activityOptions} activityLoading={activityLoading} resolvedActivity={resolvedActivity} resolvedLoading={resolvedLoading} onNext={() => { if (validateStep("business")) { setErrors({}); setStep("risk"); markReachable("risk"); } }} />}
              {step === "risk" && <StepRisk form={form} setForm={setForm} schema={schema} errors={errors} loading={uwLoading} onBack={() => setStep("business")} onNext={runAppetiteAndQuickQuote} />}
              {step === "appetite" && <StepAppetite uw={uw} onBack={() => setStep("risk")} onProceed={() => { setStep("packages"); markReachable("packages"); }} onNewQuote={resetAll} />}
              {step === "packages" && (
                <StepPackages
                  mode={mode} setMode={setMode} offers={offers} selected={selectedPkg} setSelected={setSelectedPkg}
                  flexState={{ selections: flexSelections, setSelections: setFlexSelections, priced: flexPriced }}
                  loadingFlex={flexLoading} loadingOffers={false} onFlexRecalc={runFlexBind}
                  onOpenFlex={() => { if (!flexRequirements) loadFlexRequirements(); }} flexRequirements={flexRequirements} flexReqLoading={flexReqLoading}
                  onBack={() => setStep("appetite")}
                  onContinue={(src) => { setBindSource(src || "quick"); setBindResult(src === "flex" ? flexPriced : null); setStep("bind"); markReachable("bind"); }}
                />
              )}
              {step === "bind" && (
                <StepBind bindResult={bindSource === "flex" ? flexPriced : bindResult} bindIssues={bindIssues} source={bindSource} form={form} selectedName={selectedName} selectedPkg={selectedPkg}
                  endorsement={endorsement} setEndorsement={setEndorsement} loading={bindLoading} onRun={() => runQuickBind("Re-bind")}
                  onBack={() => setStep("packages")} onContinue={() => { setStep("issue"); markReachable("issue"); }} />
              )}
              {step === "issue" && (
                <StepIssue form={form} uw={uw} bindResult={bindSource === "flex" ? flexPriced : bindResult} source={bindSource} selectedName={selectedName} quoteRef={quoteRef} onNewQuote={resetAll} />
              )}
            </FadeIn>
          </div>
        </div>
        {showTicker && <LiveTicker form={form} ticker={ticker} onCheckNow={() => runTickerCheck("Eligibility Check")} />}
      </div>

      <div style={{ textAlign: "center", padding: "18px 24px 28px", fontSize: 10.5, color: T.muted }}>
        <a href={TESTER_URL} target="_blank" rel="noreferrer" style={{ color: T.muted, textDecoration: "underline" }}>{SERVICE}</a>
        {serviceMeta?.version && <span> · v{serviceMeta.version}</span>}
      </div>

      <CalculatingOverlay visible={overlay} />
      <AIAssistant form={form} setForm={setForm} step={step} offers={offers} activityOptions={activityOptions} industryMapping={industryMapping} />
      <ApiLogTab onClick={() => setApiLogOpen(true)} />
      {apiLogOpen && <ApiLogModal log={apiLog} onClose={() => setApiLogOpen(false)} />}
    </div>
  );
}
