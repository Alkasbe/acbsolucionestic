document.getElementById("convertToExcel").addEventListener("click", convertToExcel);
document.getElementById("convertToJson").addEventListener("click", convertToJson);

function convertToExcel() {
    const fileInput = document.getElementById("fileInput").files[0];
    if (!fileInput || !fileInput.name.endsWith(".json")) {
        alert("Selecciona un archivo JSON válido.");
        return;
    }

    const fileName = getFileNameWithoutExtension(fileInput.name);

    const reader = new FileReader();
    reader.onload = function (event) {
        const jsonData = JSON.parse(event.target.result);
        const wb = XLSX.utils.book_new();

        // METADATA
        const metadata = { ...jsonData };
        delete metadata.usuarios;
        const wsMetadata = XLSX.utils.json_to_sheet([metadata], { defval: null });
        XLSX.utils.book_append_sheet(wb, wsMetadata, "Metadata");

        // USUARIOS y SERVICIOS
        const usuariosSheet = [];
        const serviciosPorTipo = {};

        jsonData.usuarios.forEach((usuario) => {
            const idUsuario = `${usuario.tipoDocumentoIdentificacion}_${usuario.numDocumentoIdentificacion}`;

            const userCopy = { ...usuario };
            delete userCopy.servicios;
            userCopy._idUsuario = idUsuario;
            usuariosSheet.push(userCopy);

            if (usuario.servicios) {
                for (const tipo in usuario.servicios) {
                    if (!serviciosPorTipo[tipo]) serviciosPorTipo[tipo] = [];
                    usuario.servicios[tipo].forEach(servicio => {
                        serviciosPorTipo[tipo].push({
                            _idUsuario: idUsuario,
                            ...servicio
                        });
                    });
                }
            }
        });

        const wsUsuarios = XLSX.utils.json_to_sheet(usuariosSheet, { defval: null });
        XLSX.utils.book_append_sheet(wb, wsUsuarios, "Usuarios");

        for (const tipo in serviciosPorTipo) {
            const ws = XLSX.utils.json_to_sheet(serviciosPorTipo[tipo], { defval: null });
            XLSX.utils.book_append_sheet(wb, ws, tipo.substring(0, 31));
        }

        XLSX.writeFile(wb, fileName + ".xlsx");
    };

    reader.readAsText(fileInput);
}

function convertToJson() {
    const fileInput = document.getElementById("fileInput").files[0];
    if (!fileInput || !fileInput.name.endsWith(".xlsx")) {
        alert("Selecciona un archivo Excel válido.");
        return;
    }

    const fileName = getFileNameWithoutExtension(fileInput.name);

    const reader = new FileReader();
    reader.onload = function (event) {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const metadata = XLSX.utils.sheet_to_json(workbook.Sheets["Metadata"], { defval: null })[0] || {};
        const usuariosData = XLSX.utils.sheet_to_json(workbook.Sheets["Usuarios"], { defval: null });

        const usuarios = usuariosData.map(u => {
            const user = { ...u };
            delete user._idUsuario;
            user.servicios = {};
            return user;
        });

        // Mapa para localizar rápidamente al usuario por su ID
        const usuariosMap = {};
        usuariosData.forEach((u, i) => {
            const id = u._idUsuario;
            if (id) usuariosMap[id] = usuarios[i];
        });

        workbook.SheetNames.forEach(sheetName => {
            if (sheetName !== "Usuarios" && sheetName !== "Metadata") {
                const registros = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
                registros.forEach(r => {
                    const idUsuario = r._idUsuario;
                    if (idUsuario && usuariosMap[idUsuario]) {
                        const servicio = { ...r };
                        delete servicio._idUsuario;

                        const usuario = usuariosMap[idUsuario];
                        if (!usuario.servicios[sheetName]) {
                            usuario.servicios[sheetName] = [];
                        }
                        usuario.servicios[sheetName].push(servicio);
                    }
                });
            }
        });

        const jsonFinal = { ...metadata, usuarios };
        const blob = new Blob([JSON.stringify(jsonFinal, null, 2)], { type: "application/json" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = fileName + ".json";
        link.click();
    };

    reader.readAsArrayBuffer(fileInput);
}

function getFileNameWithoutExtension(fileName) {
    return fileName.substring(0, fileName.lastIndexOf(".")) || fileName;
}