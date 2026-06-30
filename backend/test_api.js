import fs from "fs";
import path from "path";
import { execSync } from "child_process";

const pdfPath = "C:/Users/RehanRaza/.gemini/antigravity/brain/f9d52904-bf66-4432-8f88-f7d498085c7e/media__1782709768735.pdf";

const localPdf = path.resolve("test_amcp_guideline.pdf");
fs.copyFileSync(pdfPath, localPdf);

const config = {
  metadata: {
    brandName: "SIMU HIV-1 Quant System",
    genericName: "Simulative HIV-1 Viral Load Assay",
    manufacturer: "KadelLabs Pharma",
    versionNumber: "5.0",
    compilationDate: "April 2024"
  },
  dossierType: "format-b",
  sections: {
    clinical_evidence: localPdf
  }
};

const configPath = path.resolve("test_config.json");
fs.writeFileSync(configPath, JSON.stringify(config, null, 2));

console.log("Testing python compilation with real AMCP guidelines PDF...");
const outputPath = path.resolve("amcp_output_dossier.docx");

try {
  execSync(`python -u dossier_builder.py ${configPath} ${outputPath}`, { stdio: 'inherit' });
  console.log("Dossier compiled and validated successfully!");
} catch (error) {
  console.error("Dossier compilation failed:", error);
}

// Cleanup
if (fs.existsSync(localPdf)) fs.unlinkSync(localPdf);
if (fs.existsSync(configPath)) fs.unlinkSync(configPath);
