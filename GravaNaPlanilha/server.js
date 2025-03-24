const express = require("express");
const fs = require("fs");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(__dirname));

const FILE_PATH = "dados.xlsx";

// Função para salvar os dados na planilha
function salvarDados(nome, email) {
    let workbook;
    let worksheet;

    if (fs.existsSync(FILE_PATH)) {
        workbook = XLSX.readFile(FILE_PATH);
        worksheet = workbook.Sheets["Dados"];
    } else {
        workbook = XLSX.utils.book_new();
        worksheet = XLSX.utils.aoa_to_sheet([["Nome", "Email"]]); 
        XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
    }

    const dados = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    dados.push([nome, email]); 

    const novoWorksheet = XLSX.utils.aoa_to_sheet(dados);
    workbook.Sheets["Dados"] = novoWorksheet;
    XLSX.writeFile(workbook, FILE_PATH);
}

// Rota para salvar os dados
app.post("/salvar", (req, res) => {
    console.log("Recebido:", req.body); // Verifica se os dados estão chegando
    const { nome, email } = req.body;
    if (!nome || !email) {
        return res.status(400).json({ message: "Preencha todos os campos!" });
    }
    salvarDados(nome, email);
    res.json({ message: "Dados salvos com sucesso!" });
});


// Iniciar o servidor
app.listen(3000, () => console.log("Servidor rodando na porta 3000"));
