document.getElementById("convertToExcel").addEventListener("click", convertToExcel);
document.getElementById("convertToJson").addEventListener("click", convertToJson);

function convertToExcel() {
    let fileInput = document.getElementById('fileInput').files[0];
    if (!fileInput) {
        alert('Por favor selecciona un archivo válido desde una ubicación accesible.');
        return;
    }
    
    if (!fileInput.name.endsWith('.json')) {
        alert('Por favor selecciona un archivo JSON.');
        return;
    }
    
    let fileName = getFileNameWithoutExtension(fileInput.name);
    
    let reader = new FileReader();
    reader.onload = function(event) {
        let jsonData = JSON.parse(event.target.result);
        let wb = XLSX.utils.book_new();
        
        if (jsonData.hasOwnProperty("usuarios")) {
            let metadata = { ...jsonData };
            delete metadata.usuarios;
            let metadataSheet = [metadata];
            let wsMetadata = XLSX.utils.json_to_sheet(metadataSheet, { defval: null });
            XLSX.utils.book_append_sheet(wb, wsMetadata, "Metadata");
            
            let bloques = {};
            let usuariosSheet = [];
            jsonData.usuarios.forEach(user => {
                let userCopy = { ...user };
                delete userCopy.servicios;
                usuariosSheet.push(userCopy);
                
                if (user.servicios) {
                    for (let tipoServicio in user.servicios) {
                        if (!bloques[tipoServicio]) {
                            bloques[tipoServicio] = [];
                        }
                        user.servicios[tipoServicio].forEach(servicio => {
                            let fila = { ...servicio };
                            fila.tipoDocumentoIdentificacion = user.tipoDocumentoIdentificacion;
                            fila.numDocumentoIdentificacion = user.numDocumentoIdentificacion;
                            bloques[tipoServicio].push(fila);
                        });
                    }
                }
            });
            
            let wsUsuarios = XLSX.utils.json_to_sheet(usuariosSheet, { defval: null });
            XLSX.utils.book_append_sheet(wb, wsUsuarios, "Usuarios");
            
            for (let nombreHoja in bloques) {
                let ws = XLSX.utils.json_to_sheet(bloques[nombreHoja], { defval: null });
                XLSX.utils.book_append_sheet(wb, ws, nombreHoja.substring(0, 31));
            }
        } else if (jsonData.hasOwnProperty("ResultadosValidacion")) {
            let metadata = { ...jsonData };
            delete metadata.ResultadosValidacion;
            let metadataSheet = [metadata];
            let wsMetadata = XLSX.utils.json_to_sheet(metadataSheet, { defval: null });
            XLSX.utils.book_append_sheet(wb, wsMetadata, "Metadata");
            
            let wsResultados = XLSX.utils.json_to_sheet(jsonData.ResultadosValidacion, { defval: null });
            XLSX.utils.book_append_sheet(wb, wsResultados, "ResultadosValidacion");
        }
        
        XLSX.writeFile(wb, fileName + ".xlsx");
    };
    reader.readAsText(fileInput);
}

function convertToJson() {
    let fileInput = document.getElementById('fileInput').files[0];
    if (!fileInput) {
        alert('Por favor selecciona un archivo válido desde una ubicación accesible.');
        return;
    }
    
    if (!fileInput.name.endsWith('.xlsx')) {
        alert('Por favor selecciona un archivo Excel.');
        return;
    }
    
    let fileName = getFileNameWithoutExtension(fileInput.name);
    
    let reader = new FileReader();
    reader.onload = function(event) {
        let data = new Uint8Array(event.target.result);
        let workbook = XLSX.read(data, { type: "array" });
        let metadata = XLSX.utils.sheet_to_json(workbook.Sheets["Metadata"], { defval: null })[0] || {};
        
        if (workbook.SheetNames.includes("Usuarios")) {
            let usuarios = {};
            let usuariosData = XLSX.utils.sheet_to_json(workbook.Sheets["Usuarios"], { defval: null });
            
            usuariosData.forEach(user => {
                usuarios[user.numDocumentoIdentificacion] = { ...user, servicios: {} };
            });
            
            workbook.SheetNames.forEach(sheetName => {
                if (sheetName !== "Usuarios" && sheetName !== "Metadata") {
                    let jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
                    jsonData.forEach(row => {
                        let docKey = row.numDocumentoIdentificacion;
                        if (usuarios[docKey]) {
                            if (!usuarios[docKey].servicios[sheetName]) {
                                usuarios[docKey].servicios[sheetName] = [];
                            }
                            let newItem = { ...row };
                            delete newItem.tipoDocumentoIdentificacion;
                            delete newItem.numDocumentoIdentificacion;
                            usuarios[docKey].servicios[sheetName].push(newItem);
                        }
                    });
                }
            });
            
            let finalJson = { ...metadata, usuarios: Object.values(usuarios) };
            let blob = new Blob([JSON.stringify(finalJson, null, 2)], { type: "application/json" });
            let link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = fileName + ".json";
            link.click();
        } else if (workbook.SheetNames.includes("ResultadosValidacion")) {
            let resultadosValidacion = XLSX.utils.sheet_to_json(workbook.Sheets["ResultadosValidacion"], { defval: null });
            let finalJson = { ...metadata, ResultadosValidacion: resultadosValidacion };
            
            let blob = new Blob([JSON.stringify(finalJson, null, 2)], { type: "application/json" });
            let link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = fileName + ".json";
            link.click();
        }
    };
    reader.readAsArrayBuffer(fileInput);
}

function getFileNameWithoutExtension(fileName) {
    return fileName.substring(0, fileName.lastIndexOf(".")) || fileName;
}
