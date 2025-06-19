async function submitForm() {
    const formData = {
        companyName: document.getElementById("companyName").value,
        address: document.getElementById("address").value,
        year: document.getElementById("year").value,
        month: document.getElementById("month").value,
        day: document.getElementById("day").value,
        presidentName: document.getElementById("presidentName").value,
        presidentAddress: document.getElementById("presidentAddress").value,
        birthyear: document.getElementById("birthyear").value,
        birthmonth: document.getElementById("birthmonth").value,
        birthday: document.getElementById("birthday").value,
        purpose1: document.getElementById("purpose1").value,
        purpose2: document.getElementById("purpose2").value,
        purpose3: document.getElementById("purpose3").value,
        purpose4: document.getElementById("purpose4").value,
        purpose5: document.getElementById("purpose5").value
    };

    if (!formData.companyName || !formData.address || !formData.year || !formData.month || !formData.day || !formData.presidentName || !formData.presidentAddress || !formData.purpose1 ) {
        alert("Please fill in all required fields.");
        return;
    }

    document.getElementById("loader").style.display = "block";

    try {
        // Send data to FastAPI backend to generate the Word file
        const wordResponse = await fetch("https://onestopjpn.onrender.com/generate-word", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(formData)
        });

        const wordResponse2 = await fetch("https://onestopjpn.onrender.com/generate-word2", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(formData)
        });

        const excelResponse = await fetch("https://onestopjpn.onrender.com/generate-excel", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(formData)
        });

        if (wordResponse.ok && wordResponse2.ok && excelResponse.ok) {
            // alert("Files have been successfully generated.");
            // ダウンロードボタン
            // document.getElementById("downloadMessage").style.display = "block";
            // document.getElementById("downloadWordButton").style.display = "block";
            // document.getElementById("downloadWordButton2").style.display = "block";
            // document.getElementById("downloadExcelButton").style.display = "block";

            document.getElementById("thankYouMessage").style.display = "block";

        } else {
            alert("Error: " + (await wordResponse.json()).detail);
        }
    } catch (error) {
        alert("Submission failed: " + error.message);
    } finally {
        document.getElementById("loader").style.display = "none";
    }
}

async function downloadWordFile() {
    try {
        const response = await fetch("https://onestopjpn.onrender.com/get-created-word");

        if (!response.ok) {
            throw new Error("Failed to fetch Registration Word file");
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "created_registration.docx";
        document.body.appendChild(a);
        a.click();
        a.remove();
    } catch (error) {
        alert("Download failed: " + error.message);
    }
}

async function downloadWordFile2() {
    try {
        const response = await fetch("https://onestopjpn.onrender.com/get-created-word2");
        if (!response.ok) throw new Error("Failed to fetch Incorporation Articles Word file");
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "created_incorparticles.docx";
        document.body.appendChild(a);
        a.click();
        a.remove();
    } catch (error) {
        alert("Download failed: " + error.message);
    }
}

async function downloadExcelFile() {
    try {
        const response = await fetch("https://onestopjpn.onrender.com/get-created-excel");

        if (!response.ok) {
            throw new Error("Failed to fetch Corporation Application Excel file");
        }

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "created_corporation_application.xlsx";
        document.body.appendChild(a);
        a.click();
        a.remove();
    } catch (error) {
        alert("Download failed: " + error.message);
    }
}
