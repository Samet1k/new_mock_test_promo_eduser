const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const app = express();
const port = 3000;

const cors = require('cors');
app.use(cors());

app.use(express.json());

app.use(express.static(path.join(__dirname, 'public'))); // public директориясын көрсету

// Excel файлын оқу және пән атауына сәйкес промокод беру
app.get('/getPromoCode/:testName', (req, res) => {
    const testName = req.params.testName;
    const filePath = path.join(__dirname, 'data', 'mock_test_promo.xlsx'); // Excel файлының жолы

    // Excel файлын оқу
    const workbook = xlsx.readFile(filePath);
    const sheetNames = workbook.SheetNames; // Барлық парақтардың аттары
    let promoCodeFound = null;

    // Барлық парақтарды тексеру
    for (let sheetName of sheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        console.log(`Тест атауы: ${testName}`); // Лог
        console.log(`Бағандар: ${data[0]}`);  // Лог

        const columnIndex = data[0].indexOf(testName);
        if (columnIndex !== -1) {
            const promoCodes = data.slice(1).map(row => row[columnIndex]).filter(code => code && code !== 'Өшірілген');
            promoCodeFound = promoCodes.find(code => code); // Өшірілген емес промокодты табу

            if (promoCodeFound) {
                // Промокодты қайтару
                res.json({ promoCode: promoCodeFound });

                // Промокодты өшіру
                const promoRowIndex = data.findIndex(row => row[columnIndex] === promoCodeFound);
                data[promoRowIndex][columnIndex] = 'Өшірілген'; // Өшірілген деп белгілеу

                // Жаңартылған деректерді қайта жазу
                const updatedSheet = xlsx.utils.aoa_to_sheet(data);
                workbook.Sheets[sheetName] = updatedSheet;
                xlsx.writeFile(workbook, filePath); // Excel файлын қайта жазу
                return;
            }
        }
    }

    // Егер промокод табылмаса
    if (!promoCodeFound) {
        res.status(404).send('Промокод жоқ немесе өшірілген');
    }
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
