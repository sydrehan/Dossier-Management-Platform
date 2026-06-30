import DocxMerger from '@scholarcy/docx-merger';
import fs from 'fs';

try {
  const file1 = fs.readFileSync('./output/docx-input/cover.docx');
  const file2 = fs.readFileSync('./output/docx-input/synopsis.docx');
  const file3 = fs.readFileSync('./output/docx-input/product.docx');

  console.log('Read files. Initializing @scholarcy/docx-merger...');
  const docx = new DocxMerger({}, [file1, file2, file3]);

  docx.save('nodebuffer', (data) => {
    fs.writeFileSync("./output/test-scholarcy-output.docx", data);
    console.log("File merged successfully with @scholarcy/docx-merger! Size:", fs.statSync("./output/test-scholarcy-output.docx").size);
  });
} catch (e) {
  console.error("@scholarcy/docx-merger failed:", e);
}
