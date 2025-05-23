const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const { exec } = require('child_process'); // Для выполнения git команд
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

    try {
        // Чтение файла
        const workbook = xlsx.readFile(filePath);
        const sheetNames = workbook.SheetNames; // Барлық парақтардың аттары
        let promoCodeFound = null;

        // Барлық парақтарды тексеру
        for (let sheetName of sheetNames) {
            const sheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

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

                    // Логируем данные перед записью в файл
                    console.log("Данные перед записью в файл: ", data);

                    // Записываем обновлённый файл
                    const updatedSheet = xlsx.utils.aoa_to_sheet(data);
                    workbook.Sheets[sheetName] = updatedSheet;

                    // Записываем файл
                    const updatedFilePath = path.join(__dirname, 'data', 'mock_test_promo.xlsx');
                    xlsx.writeFile(workbook, updatedFilePath);

                    console.log("Файл успешно обновлён и сохранён в: ", updatedFilePath);

                    // Отправляем изменения в GitHub после изменения Excel
                    updateExcelFile();

                    return;
                }
            }
        }

        // Егер промокод табылмаса
        if (!promoCodeFound) {
            console.log(`Промокод "${testName}" не найден или уже был удалён.`);
            res.status(404).send('Промокод жоқ немесе өшірілген');
        }
    } catch (error) {
        console.error("Ошибка при работе с файлом Excel:", error);
        res.status(500).send('Ошибка при обработке Excel файла.');
    }
});

// Функция для отправки изменений в GitHub
const updateExcelFile = () => {
    exec('git add data/mock_test_promo.xlsx', (error, stdout, stderr) => {
        if (error) {
            console.error(`Ошибка при добавлении файла: ${error.message}`);
            return;
        }
        if (stderr) {
            console.error(`stderr: ${stderr}`);
            return;
        }

        // Делаем commit
        exec('git commit -m "Обновление промокодов в Excel файле"', (error, stdout, stderr) => {
            if (error) {
                console.error(`Ошибка при commit: ${error.message}`);
                return;
            }
            if (stderr) {
                console.error(`stderr: ${stderr}`);
                return;
            }

            // Отправляем изменения на GitHub
            exec('git push origin master', (error, stdout, stderr) => {
                if (error) {
                    console.error(`Ошибка при push: ${error.message}`);
                    return;
                }
                if (stderr) {
                    console.error(`stderr: ${stderr}`);
                    return;
                }
                console.log('Изменения успешно отправлены на GitHub');
            });
        });
    });
};

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
