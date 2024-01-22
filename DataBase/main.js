const bodyParser = require('body-parser');
const express = require('express');
const app = express();
const fs = require('fs');
const XLSX = require('xlsx');




app.set('view engine', 'ejs')
app.use(bodyParser.json());
app.use(express.static('Files'));


const workbook = XLSX.readFile("./DataBase/DataBase.xlsx")
let worksheets = {};

app.get('/getData', (req, res) => {
    for(const sheetName of workbook.SheetNames){
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }
    
    if(req.body.id && worksheets.Sheet1[req.body.id - 1]) 
    res.send(worksheets.Sheet1[req.body.id - 1]);
    else
    res.send('Возникла ошибка: вироятно был введён не верный айди))');

    
});
app.get('/getAllData', (req, res) => {
    for(const sheetName of workbook.SheetNames){
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }
    res.send(worksheets.Sheet1);
    });


app.get('/getDataVal', (req, res) => { 
    for(const sheetName of workbook.SheetNames){
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }
    res.send(worksheets.Sheet1[worksheets.Sheet1.length - 1].ID);
})

app.post('/createData', (req, res) => {
    for(const sheetName of workbook.SheetNames){
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }

    if(req.body.email && req.body.username && req.body.password){
        worksheets.Sheet1.push({
            "ID": (+worksheets.Sheet1[worksheets.Sheet1.length - 1].ID) + 1,
            "email": req.body.email,
            "username": req.body.username,
            "password": req.body.password
        });
    }else {
        res.send('Произашла ошибка, возможно у вас нету обходимых аргментов (email, username, password)')
    }

    XLSX.utils.sheet_add_json(workbook.Sheets["Sheet1"], worksheets.Sheet1);
    XLSX.writeFile(workbook, './DataBase/DataBase.xlsx');

    res.send(
        {
        "ID": (+worksheets.Sheet1[worksheets.Sheet1.length - 1].ID),
        "email": req.body.email,
        "username": req.body.username,
        "password": req.body.password
        
    });

});

app.patch('/updateData', (req, res) => {
    for (const sheetName of workbook.SheetNames) {
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }

    if(req.body.id && worksheets.Sheet1[req.body.id - 1]){
        if(req.body.email || req.body.username || req.body.password){

            if (req.body.email) {
            worksheets.Sheet1[req.body.id - 1].email = req.body.email;
            }
            if (req.body.username) {
                worksheets.Sheet1[req.body.id - 1].username = req.body.username;
            }
            if (req.body.password) {
                worksheets.Sheet1[req.body.id - 1].password = req.body.email;
            }


        }else {
            res.send('Произашла ошибка, возможно у вас нету обходимых аргментов (email, username, password)')
        }
    }else {
        res.send('Возникла ошибка: вироятно был введён не верный айди))');
    }
    
    XLSX.utils.sheet_add_json(workbook.Sheets["Sheet1"], worksheets.Sheet1);
    XLSX.writeFile(workbook, './DataBase/DataBase.xlsx');
    res.send('Данные на базе были успешно обнавлены!');


});

app.delete('/removeData', (req, res) => {
    for (const sheetName of workbook.SheetNames) {
        worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }
    if (req.body.id && worksheets.Sheet1[req.body.id - 1]) {
        
        worksheets.Sheet1.splice(req.body.id - 1)

    }
    else {
        res.status(400).send('В базе данных нет этой строки для удаления.');
    } 
    
    XLSX.utils.sheet_add_json(workbook.Sheets["Sheet1"], worksheets.Sheet1);
    XLSX.writeFile(workbook, './DataBase/DataBase.xlsx');
    res.send('Данные из базы были успешно удалены!');
})

app.use((req, res) => {
    res.status(404).render('404');
});
app.listen(8080, () => {
    console.log(`Сервер запущен: http://localhost:8080`);
})

