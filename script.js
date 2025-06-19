document.addEventListener('DOMContentLoaded', async function () {

    async function loadExcel(fileName, tableId) {
        try {
            const response = await fetch(fileName);
            const data = await response.arrayBuffer();
            const workbook = XLSX.read(data, { type: 'array' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const html = XLSX.utils.sheet_to_html(worksheet);

            let table = document.getElementById(tableId);
            table.innerHTML = html;

            table.querySelector('tr:first-child').remove();

            const rows = Array.from(table.querySelectorAll('tbody tr'));
            const newImageUrl = "https://cdn.discordapp.com/emojis/1266555790953676841.webp?size=96";

            rows.sort((a, b) => {
                const aValue = parseInt(a.querySelector('td:nth-child(4)').textContent, 10);
                const bValue = parseInt(b.querySelector('td:nth-child(4)').textContent, 10);
                return bValue - aValue;
            });

            let currentRank = 1;
            let previousValue = null;
            rows.forEach((row, index) => {
                const value = parseInt(row.querySelector('td:nth-child(4)').textContent, 10);
                if (value !== previousValue) {
                    currentRank = index + 1;
                    previousValue = value;
                }
                const rankCell = row.querySelector('td:nth-child(1)');
                if (rankCell) {
                    rankCell.textContent = currentRank;
                }

                // Style all rows
                const secondColumn = row.querySelector('td:nth-child(2)');
                if (secondColumn) {
                    const imgElement = document.createElement('img');
                    imgElement.src = newImageUrl;
                    imgElement.alt = "Image";
                    imgElement.style.width = '80px';
                    imgElement.style.height = '80px';
                    imgElement.style.display = 'block';
                    imgElement.style.margin = '0 auto';
                    secondColumn.innerHTML = '';
                    secondColumn.appendChild(imgElement);
                }

                // Apply styles to other columns
                row.querySelectorAll('td:nth-child(1), td:nth-child(3), td:nth-child(4)').forEach(cell => {
                    cell.style.flex = '1 0 100%';
                    cell.style.textAlign = 'center';
                });
            });

            rows.forEach(row => {
                if (tableId !== 'overall-tabulka') {
                    const firstColumn = row.querySelector('td:nth-child(1)');
                    if (firstColumn) {
                        firstColumn.style.display = 'none';
                    }
                    const secondColumn = row.querySelector('td:nth-child(2)');
                    if (secondColumn) {
                        secondColumn.style.display = 'none';
                    }
                } else {
                    const uuidCell = row.querySelector('td:nth-child(3)');
                    if (uuidCell) {
                        const uuid = uuidCell.textContent;
                        const imageUrl = `https://render.crafty.gg/3d/bust/${uuid}`;

                        let imageCell = row.querySelector('td:nth-child(2)');
                        if (!imageCell) {
                            imageCell = document.createElement('td');
                            row.appendChild(imageCell);
                        }

                        const imgElement = document.createElement('img');
                        imgElement.src = imageUrl;
                        imgElement.alt = `Player avatar with UUID ${uuid}`;
                        imgElement.style.width = '80px';
                        imgElement.style.height = '80px';
                        imageCell.innerHTML = '';
                        imageCell.appendChild(imgElement);

                        let combinedCell = row.querySelector('td:nth-child(5)');
                        if (!combinedCell) {
                            combinedCell = document.createElement('td');
                            row.appendChild(combinedCell);
                        }
                    }
                }
            });

            replaceNumbersInTable(tableId);
            return;
        } catch (error) {
            console.error("Error loading Excel:", error);
        }
    }

    function replaceNumbersInTable(tableId) {
        const table = document.getElementById(tableId);
        if (!table) {
            console.error("Table with ID '" + tableId + "' was not found.");
            return;
        }

        const cells = table.querySelectorAll("td");
        cells.forEach(cell => {
            const value = cell.textContent.trim();
            let newText = null;
            let textColor = null;
            let backgroundColor = null;

            switch (value) {
                case "32": newText = "HT2"; textColor = "black"; backgroundColor = "#A4B3C7"; break;
                case "16": newText = "HT3"; textColor = "black"; backgroundColor = "#8F5931"; break;
                case "10": newText = "LT3"; textColor = "black"; backgroundColor = "#B56326"; break;
                case "5": newText = "HT4"; textColor = "black"; backgroundColor = "#655B79"; break;
                case "3": newText = "LT4"; textColor = "black"; backgroundColor = "#655B79"; break;
                case "2": newText = "HT5"; textColor = "black"; backgroundColor = "#655B79"; break;
                case "1": newText = "LT5"; textColor = "black"; backgroundColor = "#655B79"; break;
                case "24": newText = "LT2"; textColor = "black"; backgroundColor = "#888D95"; break;
                case "48": newText = "LT1"; textColor = "black"; backgroundColor = "#D5B355"; break;
                case "60": newText = "HT1"; textColor = "black"; backgroundColor = "#FFCF4A"; break;
                case "22": newText = "RTL2"; textColor = "#888D95"; break;
                case "29": newText = "RHT2"; textColor = "#9EAFC6"; break;
                case "43": newText = "RTL1"; textColor = "#D5A349"; break;
                case "54": newText = "RHT1"; textColor = "#FFCC47"; break;
                default: backgroundColor = "#EEE0CB"; break;
            }

            if (newText !== null && cell.cellIndex !== 0 && cell.cellIndex !== 3) {
                cell.textContent = newText;
                cell.style.color = textColor;
                cell.style.backgroundColor = backgroundColor;

                const emojis = [
                    "1266555790953676841", "1341321180329676840", "1266550161744724060",
                    "1341321583695892575", "1266553596858732705", "1299784615149437072",
                    "1266553957543579760", "1335283642490032138"
                ];
                const emojiIndex = cell.cellIndex - 4;
                if (emojiIndex >= 0 && emojiIndex < emojis.length) {
                    const img = document.createElement('img');
                    img.src = `https://cdn.discordapp.com/emojis/${emojis[emojiIndex]}.webp?size=40`;
                    img.alt = newText;
                    img.style.display = 'block';
                    cell.appendChild(img);
                }
            }
        });
    }

    // Load data from multiple Google Sheets
    loadExcel('https://docs.google.com/spreadsheets/d/1y02Uh7eT3hEwkCrVMImjpbHLZvOmZSa3vqRPYVYTzyg/edit?usp=sharing', 'overall-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1j_F6VyWnCrt6GQxtQDjdNyOYX2h4CXR8XMwRv8Dtanw/edit?usp=sharing', 'cpvp-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1mkRA4irm2U4iWAtaM4GE-Cud3iZsaO-YU0AB8gvjhnM/edit?usp=sharing', 'axe-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1y02Uh7eT3hEwkCrVMImjpbHLZvOmZSa3vqRPYVYTzyg/edit?usp=sharing', 'sword-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1pt1KCOXspTBCEj6C6q2bnCJBLj5VJG57rXQ1vHcAJwM/edit?usp=sharing', 'npot-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/19fgMlbGaQ716KUa8umsMHk0wTZFa0leAtGpIb_44iT0/edit?usp=sharing', 'pot-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/13OqD1PetWvu7IOn6vph06m8TmML5UsCvwmaVADT-kkg/edit?usp=sharing', 'smp-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1C8Sa9pcGNzFR5gTR9lbcP9d9Wyhb9yIhx9GsCOtTfTg/edit?usp=sharing', 'uhc-tabulka');
    loadExcel('https://docs.google.com/spreadsheets/d/1AgzOlXw6C-i1rwsDs3jA3Rg2QyN_O6ZqheiHENncWsI/edit?usp=sharing', 'diasmp-tabulka');

    function showTable(tableId) {
        const allTables = document.querySelectorAll('.tabulka');
        allTables.forEach(table => table.classList.remove('active'));

        const selectedTable = document.getElementById(tableId);
        if (selectedTable) {
            selectedTable.classList.add('active');
        }
    }

    const links = document.querySelectorAll('nav a');
    links.forEach(link => {
        link.addEventListener('click', function (event) {
            event.preventDefault();
            const tableId = this.getAttribute('href').substring(1) + '-tabulka';
            showTable(tableId);
        });
    });

    showTable('overall-tabulka');
});

