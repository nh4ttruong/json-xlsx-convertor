document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("inputFile");
  const convertButton = document.getElementById("convertButton");
  const textOutput = document.getElementById("textOutput");
  const textInput = document.getElementById("textInput");
  const textAreaInput = document.getElementById("textAreaInput");
  const downloadButton = document.getElementById("downloadButton");
  const notification = document.getElementById("notification");
  const srcRadioInput = document.getElementsByName("src");
  const destRadioInput = document.getElementsByName("dest");
  const outputArea = document.getElementById("outputArea");
  const srcDestRadioInput = [...srcRadioInput, ...destRadioInput];

  srcDestRadioInput.forEach((radio) => {
    radio.addEventListener("click", function () {
      const src = document.querySelector('input[name="src"]:checked').value;
      const dest = document.querySelector('input[name="dest"]:checked').value;

      notification.classList.add("d-none");

      if (src === dest) {
        if (src === "xlsx") {
          document.getElementById("jsonDest").checked = true;
          textInput.classList.remove("d-block");
          textInput.classList.add("d-none");
          fileInput.disabled = false;
        } else if (src === "json") {
          document.getElementById("xlsxDest").checked = true;
          textInput.classList.remove("d-none");
          textInput.classList.add("d-block");
          textAreaInput.disabled = false;
          textAreaInput.value = "";
          fileInput.disabled = false;
        }
      }
    });
  });

  textAreaInput.addEventListener("input", function () {
    if (textAreaInput.value) {
      fileInput.disabled = true;
    } else {
      fileInput.disabled = false;
    }
  });

  fileInput.addEventListener("change", function () {
    // disable textInput if fileInput is selected
    if (fileInput.files[0]) {
      textAreaInput.disabled = true;
    } else {
      textAreaInput.disabled = false;
    }
    const file = fileInput.files[0];
    if (file) {
      notification.style.display = "none";
      convertButton.disabled = false;
    } else {
      addNotification("block", "Please select a file.", "bg-warning", "bg-danger");
      convertButton.disabled = true;
    }
  });

  convertButton.addEventListener("click", function () {
    const src = document.querySelector('input[name="src"]:checked').value;
    const dest = document.querySelector('input[name="dest"]:checked').value;
    const file = fileInput.files[0];

    if (src === "xlsx") {

      if (!file) {
        handleConversionError("Please select a file.");
        return;
      }

      const reader = new FileReader();

      reader.onload = function (e) {
        const result = e.target.result;
        try {
          const workbook = XLSX.read(result, { type: "binary" });
          const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
          const jsonOutput = JSON.stringify(jsonData, null, 2);
          textOutput.value = jsonOutput;

          const blob = new Blob([jsonOutput], { type: "application/json" });
          downloadButton.href = window.URL.createObjectURL(blob);

          const originalFilename = file.name + ".json";
          downloadButton.setAttribute("download", formatFilename(originalFilename) + ".json"  );
          downloadButton.classList.remove("btn-secondary");
          downloadButton.classList.add("btn-success");

          addNotification("block", "Conversion successful!", "bg-success", "bg-danger");

          outputArea.classList.remove("d-none");
          outputArea.classList.add("d-block");
        } catch (error) {
          handleConversionError("Error converting Excel to JSON. Please check the file format.");
        }
      }
      reader.readAsBinaryString(file);
    } else if (src === "json") {
      if (!textAreaInput && !file) {
        handleConversionError("Please enter JSON or select a file.");
        return;
      } else if (textAreaInput.value && file) {
        handleConversionError("Please enter JSON or select a file, not both.");
        return;
      } else if (textAreaInput.value) {
        const data = textAreaInput.value.replace(/'/g, '"').replace(/(\w+):/g, '"$1":');;
        json2excel(data);
      } else if (file) {
        const reader = new FileReader();
        reader.onload = function (event) {
          const result = event.target.result;
          json2excel(result, file);
        }
        reader.readAsText(file);
      }
    }
  });
  
  function json2excel(data, file=null) {
    try {
      console.log(data);

      if (data[0] !== "[" && data[data.length - 1] !== "]") {
        data = "[" + data + "]";
      }

      const jsonData = JSON.parse(data);
      const ws = XLSX.utils.json_to_sheet(jsonData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

      const blob = new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      });

      downloadButton.href = window.URL.createObjectURL(blob);

      const originalFilename = file ? file.name : "yourjson";
     
      downloadButton.setAttribute("download", formatFilename(originalFilename) + ".xlsx");
      textOutput.value = "";
      
      addNotification("block", "Conversion from JSON to Excel successful!", "bg-success", "bg-danger");

      downloadButton.classList.remove("btn-secondary");
      downloadButton.classList.add("btn-success");
    } catch (error) {
      handleConversionError("Error converting JSON to Excel. Please check the JSON format.");
    }
  }

  function handleConversionError(errorMessage) {
    notification.classList.add("bg-danger");
    notification.classList.remove("d-none");
    notification.classList.add("d-block");
    notification.textContent = errorMessage;
    notification.style.display = "block";
  }

  function formatFilename(filename) {
    const date = new Date();
    const year = date.getFullYear();
    const month = date.getMonth() + 1;
    const day = date.getDate();

    return `${year}-${month}-${day}_converted_${filename.replace(/\.[^.]+$/, "")}`;
  }

  function addNotification(display, message, colorAdd, colorRemove) {
    notification.classList.remove("d-none");
    notification.classList.add("d-block");
    notification.style.display = display;
    notification.textContent = message;
    notification.classList.add(colorAdd);
    notification.classList.remove(colorRemove);
    if (colorAdd === "bg-warning") {
      notification.classList.add("text-dark");
    }
   }
  
});
