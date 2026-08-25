/* =========================================================
   LOGIN
========================================================= */

function loginAlert() {
    alert("ระบบเข้าสู่ระบบจะเปิดให้ใช้งานในภายหลัง");
}


/* =========================================================
   ตัวแปร
========================================================= */

let usageChart = null;


/*
   เก็บข้อมูลกราฟวงกลม
*/
let pieData = {
    tunnel: 0,
    hangdong: 0,
    ecology: 0,
    agriculture: 0
};


/*
   เก็บรายการที่ผู้ใช้เลือก

   สามารถเลือกได้หลายตัวพร้อมกัน
   เช่น
   ["tunnel", "ecology"]
*/
let selectedPieFilters = [];


/* =========================================================
   โฟลเดอร์ Excel
========================================================= */

const EXCEL_FOLDER = "excel";


/* =========================================================
   แปลงค่าเป็นตัวเลข
========================================================= */

function toNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return 0;
    }

    if (typeof value === "number") {
        return value;
    }

    const number = parseFloat(
        String(value)
            .replace(/,/g, "")
            .trim()
    );

    return Number.isFinite(number)
        ? number
        : 0;
}


/* =========================================================
   จัดรูปแบบตัวเลข
========================================================= */

function formatNumber(value) {

    return Number(value).toLocaleString("en-US", {
        minimumFractionDigits: 3,
        maximumFractionDigits: 3
    });

}


/* =========================================================
   แสดง Loading
========================================================= */

function setLoading() {

    document.getElementById(
        "waterInValue"
    ).textContent = "กำลังโหลด...";

    document.getElementById(
        "totalSendValue"
    ).textContent = "กำลังโหลด...";

    document.getElementById(
        "waterSupplyValue"
    ).textContent = "กำลังโหลด...";

    document.getElementById(
        "ecoValue"
    ).textContent = "กำลังโหลด...";

    document.getElementById(
        "agricultureValue"
    ).textContent = "กำลังโหลด...";

}


/* =========================================================
   แสดง Error
========================================================= */

function setError() {

    document.getElementById(
        "waterInValue"
    ).textContent = "ไม่พบข้อมูล";

    document.getElementById(
        "totalSendValue"
    ).textContent = "ไม่พบข้อมูล";

    document.getElementById(
        "waterSupplyValue"
    ).textContent = "ไม่พบข้อมูล";

    document.getElementById(
        "ecoValue"
    ).textContent = "ไม่พบข้อมูล";

    document.getElementById(
        "agricultureValue"
    ).textContent = "ไม่พบข้อมูล";

}


/* =========================================================
   อ่าน Excel ตามปี
========================================================= */

async function loadYearData(year) {

    setLoading();

    const filePath =
        `${EXCEL_FOLDER}/${year}.xlsx`;

    try {

        console.log(
            "กำลังอ่านไฟล์:",
            filePath
        );

        const response =
            await fetch(filePath);


        if (!response.ok) {

            throw new Error(
                `ไม่สามารถเปิด ${filePath}`
            );

        }


        const arrayBuffer =
            await response.arrayBuffer();


        const workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array"
                }
            );


        const sheetName =
            workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[sheetName];


        const rows =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: 0
                }
            );


        console.log(
            `ข้อมูลปี ${year}:`,
            rows
        );


        if (!rows.length) {

            throw new Error(
                `ไฟล์ ${filePath} ไม่มีข้อมูล`
            );

        }


        calculateDashboard(rows);


    } catch (error) {

        console.error(error);

        setError();

        alert(
            "ไม่สามารถอ่านข้อมูล Excel ได้\n\n" +
            `ไฟล์ที่กำลังอ่าน:\n${filePath}\n\n` +
            "ตรวจสอบชื่อไฟล์และโฟลเดอร์ excel"
        );

    }

}


/* =========================================================
   รวมข้อมูลในคอลัมน์
========================================================= */

function sumColumn(rows, columnName) {

    return rows.reduce(
        (total, row) => {

            return total +
                toNumber(
                    row[columnName]
                );

        },
        0
    );

}


/* =========================================================
   คำนวณ Dashboard
========================================================= */

function calculateDashboard(rows) {


    /* -----------------------------------------
       ปริมาณน้ำเข้าคลอง
    ----------------------------------------- */

    const waterIn =
        sumColumn(
            rows,
            "เข้าคลอง /ล้าน ลบ.ม."
        );


    /* -----------------------------------------
       ปริมาณรวมส่งประปา
    ----------------------------------------- */

    const waterSupply =
        sumColumn(
            rows,
            "รวมส่งประปา / ล้าน ลบ.ม."
        );


    /* -----------------------------------------
       ส่งนิเวศ
    ----------------------------------------- */

    const ecology =
        sumColumn(
            rows,
            "ส่งนิเวศ / ล้าน ลบ.ม."
        );


    /* -----------------------------------------
       คงเหลือเกษตร
    ----------------------------------------- */

    const agriculture =
        sumColumn(
            rows,
            "คงเหลือเกษตร / ล้าน ลบ.ม."
        );


    /* -----------------------------------------
       ปริมาณส่งน้ำ
    ----------------------------------------- */

    const totalSend =
        waterSupply +
        ecology +
        agriculture;


    /* -----------------------------------------
       แสดงข้อมูล
    ----------------------------------------- */

    document.getElementById(
        "waterInValue"
    ).textContent =
        formatNumber(waterIn);


    document.getElementById(
        "totalSendValue"
    ).textContent =
        formatNumber(totalSend);


    document.getElementById(
        "waterSupplyValue"
    ).textContent =
        formatNumber(waterSupply);


    document.getElementById(
        "ecoValue"
    ).textContent =
        formatNumber(ecology);


    document.getElementById(
        "agricultureValue"
    ).textContent =
        formatNumber(agriculture);


    /* -----------------------------------------
       ข้อมูล 4 ส่วนของกราฟวงกลม
    ----------------------------------------- */

    const tunnel =
        sumColumn(
            rows,
            "ประปาอุโมง / ล้าน ลบ.ม."
        );


    const hangdong =
        sumColumn(
            rows,
            "ประปาหางดง / ล้าน ลบ.ม."
        );


    /* -----------------------------------------
       เก็บข้อมูล
    ----------------------------------------- */

    pieData = {

        tunnel: tunnel,

        hangdong: hangdong,

        ecology: ecology,

        agriculture: agriculture

    };


    /*
       เมื่อเปลี่ยนปี
       รีเซ็ตการเลือกทั้งหมด
    */

    selectedPieFilters = [];


    /* สร้างกราฟ */

    updatePieChart();

}


/* =========================================================
   คำนวณยอดรวมของกราฟ
========================================================= */

function getPieTotal() {

    return (
        pieData.tunnel +
        pieData.hangdong +
        pieData.ecology +
        pieData.agriculture
    );

}


/* =========================================================
   คำนวณ % ของแต่ละรายการ
========================================================= */

function getPiePercent(key) {

    const total =
        getPieTotal();


    if (total === 0) {
        return 0;
    }


    return (
        pieData[key] /
        total
    ) * 100;

}


/* =========================================================
   คำนวณ % รวมของรายการที่เลือก
========================================================= */

function getSelectedPercent() {

    /*
       ถ้ายังไม่ได้เลือกอะไร
       ให้แสดง 100%
    */

    if (
        selectedPieFilters.length === 0
    ) {

        return 100;

    }


    const total =
        getPieTotal();


    if (total === 0) {
        return 0;
    }


    let selectedValue = 0;


    selectedPieFilters.forEach(
        key => {

            selectedValue +=
                pieData[key];

        }
    );


    return (
        selectedValue /
        total
    ) * 100;

}


/* =========================================================
   Plugin แสดง % ตรงกลางกราฟ
========================================================= */

const centerTextPlugin = {

    id: "centerText",


    afterDraw(chart) {

        const {
            ctx,
            chartArea
        } = chart;


        if (!chartArea) {
            return;
        }


        const centerX =
            (
                chartArea.left +
                chartArea.right
            ) / 2;


        const centerY =
            (
                chartArea.top +
                chartArea.bottom
            ) / 2;


        /*
           ถ้าไม่ได้เลือกอะไร
           = 100%

           ถ้าเลือกหลายอัน
           = เปอร์เซ็นต์รวม
        */

        const percent =
            getSelectedPercent();


        const text =
            `${percent.toFixed(1)}%`;


        ctx.save();


        /*
           ตัวเลข
        */

        ctx.font =
            '800 32px "Noto Sans Thai", sans-serif';


        ctx.fillStyle =
            "#092f6b";


        ctx.textAlign =
            "center";


        ctx.textBaseline =
            "middle";


        ctx.fillText(
            text,
            centerX,
            centerY
        );


        ctx.restore();

    }

};


/* =========================================================
   สร้างกราฟวงกลม
========================================================= */

function updatePieChart() {

    const canvas =
        document.getElementById(
            "usageChart"
        );


    if (!canvas) {
        return;
    }


    /*
       ทำลายกราฟเก่า
    */

    if (usageChart) {

        usageChart.destroy();

        usageChart = null;

    }


    /* -----------------------------------------
       ข้อมูลกราฟ
    ----------------------------------------- */

    const values = [

        pieData.tunnel,

        pieData.hangdong,

        pieData.ecology,

        pieData.agriculture

    ];


    /* -----------------------------------------
       สีปกติ
    ----------------------------------------- */

    const normalColors = [

        "#1685df",

        "#16aa65",

        "#f7b735",

        "#8c70df"

    ];


    /*
       สีของกราฟ
       ถ้ามีการเลือก
       ส่วนที่ไม่ได้เลือกจะจางลง
    */

    const chartColors =
        normalColors.map(
            (color, index) => {

                const key =
                    [
                        "tunnel",
                        "hangdong",
                        "ecology",
                        "agriculture"
                    ][index];


                /*
                   ถ้าไม่ได้เลือกอะไร
                   ทุกสีแสดงปกติ
                */

                if (
                    selectedPieFilters.length === 0
                ) {

                    return color;

                }


                /*
                   ถ้ารายการนี้ถูกเลือก
                   แสดงสีเต็ม
                */

                if (
                    selectedPieFilters.includes(key)
                ) {

                    return color;

                }


                /*
                   รายการที่ไม่ได้เลือก
                   ทำให้จาง
                */

                return hexToRgba(
                    color,
                    0.18
                );

            }
        );


    /* -----------------------------------------
       สร้าง Chart
    ----------------------------------------- */

    usageChart =
        new Chart(
            canvas,
            {

                type: "doughnut",


                data: {

                    labels: [

                        "ประปาอุโมง",

                        "ประปาหางดง",

                        "ส่งนิเวศ",

                        "คงเหลือเกษตร"

                    ],


                    datasets: [{

                        data: values,


                        backgroundColor:
                            chartColors,


                        borderWidth: 3,

                        borderColor:
                            "#ffffff",


                        /*
                           ขนาดตอน hover
                        */

                        hoverOffset: 12

                    }]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio: false,


                    /*
                       รูตรงกลาง
                    */

                    cutout: "58%",


                    /*
                       Animation
                    */

                    animation: {

                        animateRotate: true,

                        animateScale: true,

                        duration: 800,

                        easing:
                            "easeOutCubic"

                    },


                    /*
                       เวลา hover
                    */

                    hover: {

                        mode: "nearest",

                        intersect: true

                    },


                    plugins: {

                        legend: {

                            display: false

                        },


                        tooltip: {

                            callbacks: {

                                label:
                                    function(context) {

                                        const key =
                                            [
                                                "tunnel",
                                                "hangdong",
                                                "ecology",
                                                "agriculture"
                                            ][
                                                context.dataIndex
                                            ];


                                        const percent =
                                            getPiePercent(
                                                key
                                            );


                                        return (
                                            " " +
                                            context.label +
                                            ": " +
                                            formatNumber(
                                                context.raw
                                            ) +
                                            " ล้าน ลบ.ม. (" +
                                            percent.toFixed(1) +
                                            "%)"
                                        );

                                    }

                            }

                        }

                    },


                    /*
                       กดที่ส่วนของกราฟโดยตรง
                       ก็เลือก Filter ได้
                    */

                    onClick:
                        function(
                            event,
                            elements
                        ) {

                            if (
                                !elements.length
                            ) {

                                return;

                            }


                            const index =
                                elements[0]
                                    .index;


                            const key =
                                [
                                    "tunnel",
                                    "hangdong",
                                    "ecology",
                                    "agriculture"
                                ][index];


                            togglePieFilter(
                                key
                            );

                        }

                },


                plugins: [

                    centerTextPlugin

                ]

            }
        );


    /*
       สร้างตัวกรอง
    */

    setupPieFilters();

}


/* =========================================================
   แปลง HEX เป็น RGBA
========================================================= */

function hexToRgba(
    hex,
    alpha
) {

    const cleanHex =
        hex.replace(
            "#",
            ""
        );


    const r =
        parseInt(
            cleanHex.substring(0, 2),
            16
        );


    const g =
        parseInt(
            cleanHex.substring(2, 4),
            16
        );


    const b =
        parseInt(
            cleanHex.substring(4, 6),
            16
        );


    return (
        `rgba(${r}, ${g}, ${b}, ${alpha})`
    );

}


/* =========================================================
   เปิด / ปิด Filter
========================================================= */

function togglePieFilter(key) {

    const index =
        selectedPieFilters.indexOf(
            key
        );


    /*
       ถ้ายังไม่ได้เลือก
       ให้เพิ่มเข้าไป
    */

    if (index === -1) {

        selectedPieFilters.push(
            key
        );

    }


    /*
       ถ้าเลือกอยู่แล้ว
       ให้เอาออก
    */

    else {

        selectedPieFilters.splice(
            index,
            1
        );

    }


    /*
       อัปเดตกราฟ
    */

    updatePieChart();


    /*
       อัปเดตหน้าตา Filter
    */

    updateLegendStyle();

}


/* =========================================================
   ตั้งค่าตัวกรอง
========================================================= */

function setupPieFilters() {

    const legend =
        document.querySelector(
            ".legend"
        );


    if (!legend) {
        return;
    }


    const items =
        legend.querySelectorAll(
            "div"
        );


    const keys = [

        "tunnel",

        "hangdong",

        "ecology",

        "agriculture"

    ];


    items.forEach(
        (item, index) => {

            if (!keys[index]) {
                return;
            }


            /*
               ทำให้รู้ว่ากดได้
            */

            item.style.cursor =
                "pointer";


            item.style.transition =
                "all 0.25s ease";


            item.style.userSelect =
                "none";


            item.setAttribute(
                "data-filter",
                keys[index]
            );


            /*
               คลิก Filter
            */

            item.onclick =
                function() {

                    const key =
                        this.getAttribute(
                            "data-filter"
                        );


                    togglePieFilter(
                        key
                    );

                };

        }
    );


    /*
       อัปเดตหน้าตา
    */

    updateLegendStyle();

}


/* =========================================================
   ทำให้ Filter ที่เลือกเด่น
========================================================= */

function updateLegendStyle() {

    const legend =
        document.querySelector(
            ".legend"
        );


    if (!legend) {
        return;
    }


    const items =
        legend.querySelectorAll(
            "div"
        );


    items.forEach(
        item => {

            const key =
                item.getAttribute(
                    "data-filter"
                );


            /*
               ไม่มีการเลือก
               ทุกอันปกติ
            */

            if (
                selectedPieFilters.length === 0
            ) {

                item.style.opacity =
                    "1";

                item.style.fontWeight =
                    "400";

                item.style.transform =
                    "translateX(0)";

                item.style.background =
                    "transparent";

                item.style.borderRadius =
                    "8px";

                item.style.padding =
                    "4px 8px";

                return;

            }


            /*
               ตัวที่เลือก
               เด่นขึ้น
            */

            if (
                selectedPieFilters.includes(
                    key
                )
            ) {

                item.style.opacity =
                    "1";

                item.style.fontWeight =
                    "700";

                item.style.transform =
                    "translateX(6px)";

                item.style.background =
                    "rgba(7, 91, 232, 0.08)";

                item.style.borderRadius =
                    "8px";

                item.style.padding =
                    "4px 8px";

            }


            /*
               ตัวที่ไม่ได้เลือก
               จางลง
            */

            else {

                item.style.opacity =
                    "0.35";

                item.style.fontWeight =
                    "400";

                item.style.transform =
                    "translateX(0)";

                item.style.background =
                    "transparent";

                item.style.borderRadius =
                    "8px";

                item.style.padding =
                    "4px 8px";

            }

        }
    );

}


/* =========================================================
   กราฟพยากรณ์
========================================================= */

function createEmptyForecastChart() {

    const canvas =
        document.getElementById(
            "waterChart"
        );


    if (!canvas) {
        return;
    }


    new Chart(
        canvas,
        {

            type: "line",


            data: {

                labels: [],

                datasets: []

            },


            options: {

                responsive: true,

                maintainAspectRatio: false,


                plugins: {

                    legend: {

                        display: false

                    }

                }

            }

        }
    );

}


/* =========================================================
   สร้างรายการปี 2543 - 2568
========================================================= */

function createYearOptions() {

    const yearSelect =
        document.getElementById(
            "yearSelect"
        );


    if (!yearSelect) {
        return;
    }


    /*
       ล้างตัวเลือกเดิม
    */

    yearSelect.innerHTML = "";


    /*
       สร้างปี 2568 -> 2543
    */

    for (
        let year = 2568;
        year >= 2543;
        year--
    ) {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            year;


        option.textContent =
            `ปี ${year}`;


        yearSelect.appendChild(
            option
        );

    }


    /*
       ค่าเริ่มต้น
    */

    yearSelect.value =
        "2568";

}


/* =========================================================
   เริ่มทำงาน
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function() {

        const yearSelect =
            document.getElementById(
                "yearSelect"
            );


        if (!yearSelect) {
            return;
        }


        /*
           สร้างรายการปี
        */

        createYearOptions();


        /*
           โหลดข้อมูลปีแรก
        */

        loadYearData(
            yearSelect.value
        );


        /*
           เปลี่ยนปี
        */

        yearSelect.addEventListener(
            "change",
            function() {

                loadYearData(
                    this.value
                );

            }
        );


        /*
           กราฟพยากรณ์
        */

        createEmptyForecastChart();

    }
);