const axios = require('axios');
const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');

var excelFilePath = path.join(`${__dirname}/excel-result`, 's1-manajemen.xlsx');

require('dotenv').config()

var baseURL = process.env.BASE_URL;
var origin = process.env.HEADER_ORIGIN;

function sleep(millis) {
    return new Promise(resolve => setTimeout(resolve, millis));
}

(async () => {
    let offset = 0;
    let maxItem = 10670;
    while (offset < maxItem) {
        try {
            const response = await axios.get(`${baseURL}${offset}`, {
                headers: {
                    Origin: origin
                }
            });

            if (response.data.status == 200) {
                const {
                    data
                } = response.data.data

                append_excel(data, maxItem);

                // await sleep(500);
                offset += 10;
            }
        } catch (error) {
            throw new Error(error.message)
        }
    }
})();

function rupiah(number) {
    return new Intl.NumberFormat("id-ID", {
        style: "currency",
        currency: "IDR"
    }).format(number);
}


function append_excel(data, maxItem) {
    let row = [
        [
            'No',
            'Jabatan',
            'Instansi',
            'Unit Kerja',
            'Formasi',
            '(PPPK)Khusus disabilitas? (CPNS)Dapat Diisi Disabilitas?',
            'Penghasilan(juta)',
            'Jumlah Kebutuhan',
            'Detail Informasi'
        ],
    ];

    if (fs.existsSync(excelFilePath)) {
        const workbook = xlsx.parse(fs.readFileSync(excelFilePath));
        const sheetData = workbook[0].data; // Assuming single sheet

        // Append the new data (excluding the header if not needed)
        row = sheetData.concat(
            data.map((item, index) => [sheetData.length + index, item.jabatan_nm,
                item.ins_nm,
                item.lokasi_nm,
                `${item.jp_nama} ${item.formasi_nm}`,
                item.disable ? 'Ya' : 'Tidak',
                (`${rupiah(item.gaji_min)} - ${rupiah(item.gaji_max)}`),
                item.jumlah_formasi,
                `https://sscasn.bkn.go.id/detailformasi/${item.formasi_id}`
            ])
        );
        console.log(`success requesting data from URL : ${origin} ${sheetData.length}/${maxItem}`);
    } else {
        // If the file does not exist, add header row first
        row = row.concat(
            data.map((item, index) => [index + 1, item.jabatan_nm,
                item.ins_nm,
                item.lokasi_nm,
                `${item.jp_nama} ${item.formasi_nm}`,
                item.disable ? 'Ya' : 'Tidak',
                (`${rupiah(item.gaji_min)} - ${rupiah(item.gaji_max)}`),
                item.jumlah_formasi,
                `https://sscasn.bkn.go.id/detailformasi/${item.formasi_id}`
            ])
        );
    }

    // Build the Excel buffer
    var buffer = xlsx.build([{
        name: 'Sheet1',
        data: row
    }]);

    // Save the buffer to an .xlsx file
    fs.writeFileSync(excelFilePath, buffer, 'binary');
}