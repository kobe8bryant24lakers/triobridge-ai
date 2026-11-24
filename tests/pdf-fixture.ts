/** Minimal, fully synthetic one-page PDF. Offsets are generated to exercise a
 * real production PDF worker, without checking private or binary fixtures in. */
export function syntheticPdf(): Buffer {
  const stream =
    'BT /F1 16 Tf 50 740 Td (Synthetic survey: 120 responses, 88 percent satisfied.) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let output = '%PDF-1.7\n';
  const offsets = objects.map((object, i) => {
    const offset = output.length;
    output += `${i + 1} 0 obj\n${object}\nendobj\n`;
    return offset;
  });
  const xref = output.length;
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  output += offsets.map((n) => `${String(n).padStart(10, '0')} 00000 n \n`).join('');
  output += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(output);
}
