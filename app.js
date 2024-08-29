import { readFile, writeFile, appendFile } from "fs/promises";
import XLSX from "xlsx"

async function asyncReadFile(filename) {
    try {
        let contents = await readFile(filename, 'utf-8');
        let result = [];
        
        contents = contents.split(/[\r\n]/);
        for(let i=0; i<contents.length;){
            if(contents[i].indexOf('ROOMING LIST FOR CHINASKY TOUR') !== -1 )
                i+=8;
            else if(contents[i].indexOf('房    號') !== -1)
                i+=12;
            else if(contents[i].indexOf('MR') !== -1 || contents[i].indexOf('MS') !== -1 || contents[i].indexOf('CHD') !== -1){
                let temp_array = [], concat_array = [];
                temp_array.push(contents[i-1])
                for(let j=0; temp_array.length != 8; j++){
                    if(!onlyWhitespace(contents[i+j])){
                        if(temp_array.length == 2 || temp_array.length == 5){
                            concat_array = contents[i+j].split(/(\s+)/);
                            for(let x=0; x<concat_array.length; x++){
                                if(temp_array.length == 2 && x == 0){
                                    temp_array.push(concat_array[x]+concat_array[x+2])
                                }
                                else if(!(onlyWhitespace(concat_array[x]) || (temp_array.length == 3  && x == 2)))
                                    temp_array.push(concat_array[x])
                            }
                        }else{
                            temp_array.push(contents[i+j])
                        }
                    }else if(temp_array.length == 7){
                        temp_array.push('')
                    }
                }
                i+=9;
                result.push(temp_array);
            }else{
                i++;
            }
        }

        // console.log(result)

        var ws = XLSX.utils.aoa_to_sheet(result);
        var wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "工作表");
        XLSX.writeFile(wb, "./output/result.xlsx");
    
        return;
    } catch (err) {
        console.log(err);
    }
}

function onlyLetters(str) {
    return /^[a-zA-Z]+$/.test(str);
}

function onlyWhitespace(str) {
    return !str.replace(/\s/g, '').length;
}

  
asyncReadFile('./input/file.txt');