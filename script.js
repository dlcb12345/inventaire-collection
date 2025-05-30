let generatedWorkbook = null;

document.getElementById('generate').addEventListener('click', handleGenerate);
document.getElementById('download').addEventListener('click', handleDownload);
document.getElementById('print').addEventListener('click', handlePrint);

function handleGenerate() {
  const fileInput = document.getElementById('upload');
  const mode = document.getElementById('mode-select').value;

  if (!fileInput.files.length) {
    alert("Veuillez sélectionner un fichier Excel.");
    return;
  }

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(sheet);

    generatedWorkbook = XLSX.utils.book_new();

    if (mode === 'global') {
      const sheet = XLSX.utils.json_to_sheet(jsonData);
      XLSX.utils.book_append_sheet(generatedWorkbook, sheet, 'Liste globale');
    } else {
      const grouped = {};
      jsonData.forEach(row => {
        const key = row["Index utilisateur"] || "Autre";
        if (!grouped[key]) grouped[key] = [];
        grouped[key].push(row);
      });

      for (let key in grouped) {
        const sheet = XLSX.utils.json_to_sheet(grouped[key]);
        XLSX.utils.book_append_sheet(generatedWorkbook, sheet, key);
      }
    }

    document.getElementById('status').innerText = "Fichier généré avec succès.";
    document.getElementById('download').disabled = false;
    document.getElementById('print').disabled = false;
  };

  reader.readAsArrayBuffer(fileInput.files[0]);
}

function handleDownload() {
  if (!generatedWorkbook) return;
  XLSX.writeFile(generatedWorkbook, 'inventaire.xlsx');
}

function handlePrint() {
  alert("Pour imprimer, veuillez d'abord télécharger le fichier Excel, l'ouvrir, puis imprimer depuis Excel.");
}
