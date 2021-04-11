
const { exec } = require("child_process");
const path = require('path');

const  xl = require('excel4node');

const csv = require('csv-parser');
const fs = require('fs');


const wb = new xl.Workbook();
const ws = wb.addWorksheet('Sheet 1');

const regex = /(.*.csv)/g;
const regex2 = /^\d+$/g;

let key_array = [];
let value_array = [];



function csv_filter(value) {

  const csv_only = value.match(regex);
  return csv_only ;

};

function convert(file_path, file_name){

  return new Promise((resolve,reject)=>{

    fs.createReadStream(file_path)
    .pipe(csv())
      
    .on('data', (row) => {
    
      for (const [key] of Object.entries(row)) {
      
        key_array.push(key);
      
      }
    
      value_array.push(Object.values(row));
    
    })

    .on('end', () => {

      const header = [ ...new Set(key_array)];  

      let k = 1; 

      for(let i=0;i<header.length;i++){

        ws.cell(1,k).string(header[i])

        let c = 2;

        for(let a=0;a<value_array.length;a++){

          (value_array[a][i].match(regex2)) ? ws.cell(c,k).number(Number(value_array[a][i])) : ws.cell(c,k).string(value_array[a][i]);
          c++;

        }

        k++;

      }


      wb.write(`Parsed/${file_name[0]}.xlsx`);
      key_array = [];
      value_array = [];
      resolve(`Parsed/${file_name[0]}.xlsx DONE`);

    });

  });

};



const wd = new Promise((resolve,reject)=>{

  const command = "echo %cd%";

  const pwd = function (command,callback){
    exec(command, (error,stdout, stderr)=>{
      callback(stdout)
    });
  };

  pwd(command,(callback)=>{

    const arr = callback.split("\r");
    arr.pop();
    resolve(arr);

  });

});


async function Do(){

  await wd.then((dir)=>{
    
    exec("dir /b", async (error, stdout, stderr) => {
      if (error) {
        console.log(`error: ${error.message}`);
        return;
      }
      if (stderr) {
        console.log(`stderr: ${stderr}`);
        return;
      }
      
      exec("md Parsed"),(err, out, stder) => {
        if (err) {
          console.log(`error: ${err.message}`);
          return;
        }
        if (stder) {
          console.log(`stderr: ${stder}`);
          return;
        }
        console.log(`Creare una cartella: ${out}`);
      }
    
      const files = stdout.split("\n");
      const filtered = files.filter(csv_filter);

    
      for(let d=0; d<filtered.length;d++){
    
        const file_name = filtered[d].split('.');
        
        const file_path = path.join(dir[0],filtered[d]).split('\r');
        console.log(file_path[0]);

        await convert(file_path[0], file_name).then(res=>console.log(res)).catch(err=>console.log(err));

    
      };

      setTimeout(()=>{},3000);
      
    });

  });

  
};

Do();






