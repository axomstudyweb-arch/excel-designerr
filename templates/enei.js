function performPDFExport() {
    if (!IS_FILE_LOADED) {
        return showCustomAlert("Upload Excel file first");
    }
    const element = document.getElementById('Ank-Viewport-Wrapper');
      const options = {
    margin:       [0.5, 0.5, 0.5, 0.5], // Margins in inches
    filename:     'Spreadsheet.pdf',
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true },
    jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
  };

  // Generate and download PDF
  html2pdf().set(options).from(element).save();
}
/**tried */
function generatePDF() {
  // Target the element you want to turn into a PDF (e.g., your spreadsheet container)
  const element = document.getElementById('spreadsheet-container'); 

  const options = {
    margin:       [0.5, 0.5, 0.5, 0.5], // Margins in inches
    filename:     'Spreadsheet.pdf',
    image:        { type: 'jpeg', quality: 0.98 },
    html2canvas:  { scale: 2, useCORS: true },
    jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
  };

  // Generate and download PDF
  html2pdf().set(options).from(element).save();
}