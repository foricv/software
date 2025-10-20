// ===============================
// Start PDF generation
// ===============================
function startPDF() {
  const btn = document.getElementById("genBtn");
  const logBox = document.getElementById("logBox");
  btn.disabled = true;
  logBox.innerHTML = "<div class='info'>⚙️ Starting PDF generation...</div>";

  fetch("/generate-pdf")
    .then((r) => r.json())
    .then((d) => {
      if (d.status === "started") {
        watchLogs();
      } else {
        logBox.innerHTML += "<div class='error'>❌ Already running or error.</div>";
        btn.disabled = false;
      }
    });
}

// ===============================
// Watch Flask log stream (Server Sent Events)
// ===============================
function watchLogs() {
  const logBox = document.getElementById("logBox");
  const evtSource = new EventSource("/logs");

  evtSource.onmessage = function (e) {
    const line = e.data;
    if (line.includes("✅"))
      logBox.innerHTML += "<div class='success'>" + line + "</div>";
    else if (line.includes("❌") || line.includes("⚠️"))
      logBox.innerHTML += "<div class='error'>" + line + "</div>";
    else
      logBox.innerHTML += "<div class='info'>" + line + "</div>";

    logBox.scrollTop = logBox.scrollHeight;

    if (line.includes("🚀 Complete process finished successfully!")) {
      evtSource.close();
      document.getElementById("genBtn").disabled = false;
    }
  };
}

// ===============================
// CNIC Auto Formatter
// ===============================
document.addEventListener("DOMContentLoaded", function () {
  const cnicInput = document.getElementById("CNIC");
  if (cnicInput) {
    cnicInput.addEventListener("input", function (e) {
      let value = e.target.value.replace(/\D/g, "");
      let formatted = "";
      if (value.length <= 5) {
        formatted = value;
      } else if (value.length <= 12) {
        formatted = value.slice(0, 5) + "-" + value.slice(5);
      } else {
        formatted = value.slice(0, 5) + "-" + value.slice(5, 12) + "-" + value.slice(12, 13);
      }
      e.target.value = formatted;
    });
  }
});

// ===============================
// Populate Dropdowns by Country → Type → Company → Project
// with fallback if "Type" not selected
// ===============================
document.addEventListener("DOMContentLoaded", function () {
  const autoRadio = document.getElementById("mode_auto");
  const manualRadio = document.getElementById("mode_manual");
  const manualBlock = document.getElementById("manual_experiences");

  const exp1 = {
    country: document.getElementById("exp1_country"),
    type: document.getElementById("exp1_type"),
    company: document.getElementById("exp1_company"),
    project: document.getElementById("exp1_project"),
  };

  const exp2 = {
    country: document.getElementById("exp2_country"),
    type: document.getElementById("exp2_type"),
    company: document.getElementById("exp2_company"),
    project: document.getElementById("exp2_project"),
  };

  function toggleMode() {
    manualBlock.style.display = manualRadio.checked ? "block" : "none";
  }
  if (autoRadio && manualRadio) {
    autoRadio.addEventListener("change", toggleMode);
    manualRadio.addEventListener("change", toggleMode);
    toggleMode();
  }

  fetch("/exp-samples")
    .then((r) => r.json())
    .then((data) => {
      if (data.status !== "ok") {
        console.error("Error fetching exp-samples:", data.message);
        return;
      }

      const rows = data.rows;

      // Build nested structure: Country → Type → Company → Projects[]
      const dataMap = {};
      rows.forEach((r) => {
        const country = r.Country.trim();
        const type = r["Company Type"].trim();
        const company = r.Company.trim();
        const project = r.Project.trim();

        if (!dataMap[country]) dataMap[country] = {};
        if (!dataMap[country][type]) dataMap[country][type] = {};
        if (!dataMap[country][type][company]) dataMap[country][type][company] = [];
        if (project && !dataMap[country][type][company].includes(project)) {
          dataMap[country][type][company].push(project);
        }
      });

      const countries = Object.keys(dataMap);

      function fillSelect(selectEl, items) {
        selectEl.innerHTML = "<option value=''>-- Select --</option>";
        items.forEach((item) => {
          const opt = document.createElement("option");
          opt.value = item;
          opt.textContent = item;
          selectEl.appendChild(opt);
        });
      }

      function getAllCompanies(country) {
        const typeData = dataMap[country] || {};
        const allCompanies = [];
        Object.keys(typeData).forEach((type) => {
          Object.keys(typeData[type]).forEach((company) => {
            if (!allCompanies.includes(company)) allCompanies.push(company);
          });
        });
        return allCompanies;
      }

      function setup(exp) {
        fillSelect(exp.country, countries);

        exp.country.addEventListener("change", function () {
          const selectedCountry = this.value;
          const types = Object.keys(dataMap[selectedCountry] || {});
          fillSelect(exp.type, types);

          // ✅ Auto-populate all companies even if no type selected
          const allCompanies = getAllCompanies(selectedCountry);
          fillSelect(exp.company, allCompanies);
          fillSelect(exp.project, []);
        });

        exp.type.addEventListener("change", function () {
          const selectedCountry = exp.country.value;
          const selectedType = this.value;

          // If no type selected → show all companies of that country
          let companies = [];
          if (!selectedType) {
            companies = getAllCompanies(selectedCountry);
          } else {
            companies = Object.keys(dataMap[selectedCountry]?.[selectedType] || {});
          }
          fillSelect(exp.company, companies);
          fillSelect(exp.project, []);
        });

        exp.company.addEventListener("change", function () {
          const selectedCountry = exp.country.value;
          const selectedType = exp.type.value;
          const selectedCompany = this.value;

          let projects = [];

          if (!selectedType) {
            // If type not selected, find company in any type
            const allTypes = Object.keys(dataMap[selectedCountry] || {});
            allTypes.forEach((t) => {
              const projList = dataMap[selectedCountry]?.[t]?.[selectedCompany];
              if (projList) projects.push(...projList);
            });
            projects = [...new Set(projects)];
          } else {
            projects = dataMap[selectedCountry]?.[selectedType]?.[selectedCompany] || [];
          }

          fillSelect(exp.project, projects);
        });
      }

      setup(exp1);
      setup(exp2);
    })
    .catch((err) => console.error("Fetch error:", err));
});

// ===============================
// Clear All Data (MainData.xlsx)
// ===============================
document.addEventListener("DOMContentLoaded", function () {
  const clearBtn = document.getElementById("clearBtn");
  if (clearBtn) {
    clearBtn.addEventListener("click", function () {
      if (!confirm("⚠️ Are you sure you want to clear all data from MainData.xlsx?")) return;

      fetch("/clear-data", { method: "POST" })
        .then((r) => r.json())
        .then((data) => {
          if (data.status === "ok") {
            alert("✅ " + data.message);
            location.reload(); // refresh page to show empty table
          } else {
            alert("❌ Error: " + data.message);
          }
        })
        .catch((err) => {
          alert("❌ Request failed: " + err);
        });
    });
  }
});
// for proper font
document.addEventListener("DOMContentLoaded", function () {
  const fieldsToFormat = ['Name', 'Designation', 'fname', 'Domicile', 'Address'];

  fieldsToFormat.forEach(function (fieldId) {
    const input = document.getElementById(fieldId);

    if (input) {
      input.addEventListener("blur", function () {
        this.value = toProperCase(this.value);
      });
    }
  });

  function toProperCase(text) {
    return text
      .toLowerCase()
      .split(' ')
      .filter(word => word.length > 0)
      .map(word => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }
});
