/**
 * cumilla_result.js - Dedicated Script for Cumilla Education Board Results
 * Direct portal: Roll Number ONLY (No Registration and No Captcha required).
 * Supports Single Roll lookup, Batch File upload, and Direct Paste Rolls.
 */

document.addEventListener("DOMContentLoaded", () => {
    // Exam and Year inputs
    const examSelect = document.getElementById("exam");
    const yearSelect = document.getElementById("year");
    const board = "comilla";

    // Single student check elements
    const cumillaSingleForm = document.getElementById("cumillaSingleForm");
    const indRollInput = document.getElementById("indRoll");
    const submitIndBtn = document.getElementById("submitIndBtn");

    // Marksheet elements
    const individualResultCard = document.getElementById("individualResultCard");
    const marksheetName = document.getElementById("marksheetName");
    const marksheetExam = document.getElementById("marksheetExam");
    const marksheetYear = document.getElementById("marksheetYear");
    const marksheetBoard = document.getElementById("marksheetBoard");
    const marksheetGpa = document.getElementById("marksheetGpa");
    const marksheetResultBadge = document.getElementById("marksheetResultBadge");
    const marksheetRoll = document.getElementById("marksheetRoll");
    const marksheetReg = document.getElementById("marksheetReg");
    const marksheetGroup = document.getElementById("marksheetGroup");
    const marksheetType = document.getElementById("marksheetType");
    const marksheetFather = document.getElementById("marksheetFather");
    const marksheetMother = document.getElementById("marksheetMother");
    const marksheetInstitute = document.getElementById("marksheetInstitute");
    const marksheetGradesBody = document.getElementById("marksheetGradesBody");

    // Batch File upload elements
    const cumillaBatchFileForm = document.getElementById("cumillaBatchFileForm");
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const fileInfo = document.getElementById("fileInfo");
    const fileName = document.getElementById("fileName");
    const demoModeCheck = document.getElementById("demoMode");
    const submitBulkBtn = document.getElementById("submitBulkBtn");

    // Paste Rolls elements
    const cumillaPasteForm = document.getElementById("cumillaPasteForm");
    const pastedRollsInput = document.getElementById("pastedRolls");
    const pasteDemoModeCheck = document.getElementById("pasteDemoMode");
    const submitPasteBtn = document.getElementById("submitPasteBtn");

    // Progress elements
    const progressCard = document.getElementById("progressCard");
    const progressTitle = document.getElementById("progressTitle");
    const progressBar = document.getElementById("progressBar");
    const progressPercentBadge = document.getElementById("progressPercentBadge");
    const statusMessage = document.getElementById("statusMessage");
    const studentCounter = document.getElementById("studentCounter");

    // Download & Navigation elements
    const downloadSection = document.getElementById("downloadSection");
    const downloadXlsxBtn = document.getElementById("downloadXlsxBtn");
    const downloadCsvBtn = document.getElementById("downloadCsvBtn");
    const openAnalysisBtn = document.getElementById("openAnalysisBtn");

    // Results table elements
    const resultsBody = document.getElementById("resultsBody");
    const resultsCount = document.getElementById("resultsCount");

    let eventSource = null;
    let extractedCount = 0;

    // ==========================================
    // KNOWN INSTITUTIONS & CUMILLA DEFAULTS
    // ==========================================
    const KNOWN_INSTITUTIONS = {
        "132092": {
            name: "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE",
            address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "132092"
        },
        "105820": {
            name: "COMILLA AJIT GUHA COLLEGE",
            address: "Chawkbazar, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "105820"
        },
        "105814": {
            name: "CUMILLA VICTORIA GOVT. COLLEGE",
            address: "Victoria College Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "105814"
        },
        "105786": {
            name: "CUMILLA ZILLA SCHOOL",
            address: "Kandirpar, Cumilla Sadar, Cumilla, Bangladesh",
            eiin: "105786"
        },
        "106579": {
            name: "FENI GOVT. COLLEGE",
            address: "College Road, Feni - 3900, Bangladesh",
            eiin: "106579"
        },
        "103284": {
            name: "BRAHMANBARIA GOVT. COLLEGE",
            address: "College Road, Brahmanbaria - 3400, Bangladesh",
            eiin: "103284"
        },
        "103597": {
            name: "CHANDPUR GOVT. COLLEGE",
            address: "Stadium Road, Chandpur - 3600, Bangladesh",
            eiin: "103597"
        },
        "107541": {
            name: "NOAKHALI GOVT. COLLEGE",
            address: "Maijdee Court, Noakhali - 3800, Bangladesh",
            eiin: "107541"
        },
        "107386": {
            name: "BANDERHAT HIGH SCHOOL & COLLEGE",
            address: "Banderhat, Lakshmipur, Bangladesh",
            eiin: "107386"
        },
        "107977": {
            name: "DHAKA COLLEGE",
            address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh",
            eiin: "107977"
        }
    };

    function normalizeInstKey(name) {
        return (name || "").toString().toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
    }

    function getSavedEiinOverride(instName) {
        try {
            const key = normalizeInstKey(instName);
            if (!key) return "";
            return localStorage.getItem(`eiin_override::${key}`) || "";
        } catch (e) { return ""; }
    }

    function saveEiinOverride(instName, eiin) {
        try {
            const key = normalizeInstKey(instName);
            if (!key) return;
            if (eiin && /^\d{4,8}$/.test(eiin)) {
                localStorage.setItem(`eiin_override::${key}`, eiin);
            }
        } catch (e) {}
    }

    function resolveInstitutionInfo(rawInstitute, rawEiin, rawAddress) {
        let eiin = (rawEiin || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(eiin.toLowerCase())) eiin = "";
        let instName = (rawInstitute || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(instName.toLowerCase())) instName = "";
        let address = (rawAddress || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(address.toLowerCase())) address = "";

        if (instName) {
            const m = instName.match(/\(?\b(\d{5,7})\b\)?/);
            if (m) {
                if (!eiin) eiin = m[1];
                instName = instName.replace(/\(?\b\d{5,7}\b\)?/g, "").replace(/[-()]/g, " ").replace(/\s+/g, " ").trim();
            }
        }

        if (!eiin && instName) {
            eiin = getSavedEiinOverride(instName);
        }

        if (eiin && KNOWN_INSTITUTIONS[eiin]) {
            const match = KNOWN_INSTITUTIONS[eiin];
            return {
                name: instName || match.name,
                eiin: eiin,
                address: address || match.address
            };
        }

        if (!eiin && instName) {
            const upper = instName.toUpperCase();
            for (const [kEiin, data] of Object.entries(KNOWN_INSTITUTIONS)) {
                if (upper === data.name || upper.includes(data.name)) {
                    eiin = kEiin;
                    if (!address) address = data.address;
                    break;
                }
            }
        }

        const cumillaDistricts = ["Cumilla", "Comilla", "Feni", "Brahmanbaria", "Chandpur", "Noakhali", "Lakshmipur", "Laxmipur"];
        if (!address && instName) {
            for (const dist of cumillaDistricts) {
                if (new RegExp(`\\b${dist}\\b`, "i").test(instName)) {
                    const disp = dist.toLowerCase() === "comilla" ? "Cumilla" : (dist.toLowerCase() === "laxmipur" ? "Lakshmipur" : dist);
                    address = `${disp}, Bangladesh`;
                    break;
                }
            }
        }

        if (!instName) {
            instName = "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE";
            if (!eiin) eiin = "132092";
            if (!address) address = "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh";
        } else if (!address) {
            address = "Cumilla, Bangladesh";
        }

        return {
            name: instName,
            eiin: eiin,
            address: address
        };
    }

    function updateMarksheetInstitutionalHeader(info) {
        const marksheetHeaderInstName = document.getElementById("marksheetHeaderInstName");
        const marksheetHeaderAddress = document.getElementById("marksheetHeaderAddress");
        const marksheetHeaderEiin = document.getElementById("marksheetHeaderEiin");
        if (marksheetHeaderInstName) marksheetHeaderInstName.textContent = info.name || "EDUCATIONAL INSTITUTION";
        if (marksheetHeaderAddress) marksheetHeaderAddress.textContent = info.address || "Cumilla, Bangladesh";
        if (marksheetHeaderEiin) marksheetHeaderEiin.textContent = info.eiin || "N/A";
        try {
            if (info.eiin) {
                sessionStorage.setItem("last_eiin", info.eiin);
            } else {
                sessionStorage.removeItem("last_eiin");
            }
            sessionStorage.setItem("last_inst", info.name || "");
            sessionStorage.setItem("last_address", info.address || "");
        } catch (e) {}
    }

    // Dynamic listener on EIIN editing: queries full Bangladesh EIIN database via API
    const marksheetHeaderEiinEl = document.getElementById("marksheetHeaderEiin");
    let eiinLookupTimer = null;
    if (marksheetHeaderEiinEl) {
        marksheetHeaderEiinEl.addEventListener("input", () => {
            const typedEiin = marksheetHeaderEiinEl.textContent.trim();
            const currentInst = document.getElementById("marksheetHeaderInstName")?.textContent?.trim() || "";
            if (/^\d{4,8}$/.test(typedEiin) && currentInst) {
                saveEiinOverride(currentInst, typedEiin);
            }
            if (typedEiin.length >= 5 && /^\d{5,7}$/.test(typedEiin)) {
                if (KNOWN_INSTITUTIONS[typedEiin]) {
                    const match = KNOWN_INSTITUTIONS[typedEiin];
                    const marksheetHeaderInstName = document.getElementById("marksheetHeaderInstName");
                    const marksheetHeaderAddress = document.getElementById("marksheetHeaderAddress");
                    if (marksheetHeaderInstName) marksheetHeaderInstName.textContent = match.name;
                    if (marksheetHeaderAddress) marksheetHeaderAddress.textContent = match.address;
                    try {
                        sessionStorage.setItem("last_eiin", match.eiin);
                        sessionStorage.setItem("last_inst", match.name);
                        sessionStorage.setItem("last_address", match.address);
                    } catch (e) {}
                    return;
                }
                clearTimeout(eiinLookupTimer);
                eiinLookupTimer = setTimeout(async () => {
                    try {
                        const resp = await fetch(`/api/lookup-institution?eiin=${encodeURIComponent(typedEiin)}`);
                        if (!resp.ok) return;
                        const data = await resp.json();
                        if (data.found && data.name) {
                            const marksheetHeaderInstName = document.getElementById("marksheetHeaderInstName");
                            const marksheetHeaderAddress = document.getElementById("marksheetHeaderAddress");
                            if (marksheetHeaderInstName) marksheetHeaderInstName.textContent = data.name;
                            if (marksheetHeaderAddress && data.address) marksheetHeaderAddress.textContent = data.address;
                            try {
                                sessionStorage.setItem("last_eiin", data.eiin || typedEiin);
                                sessionStorage.setItem("last_inst", data.name);
                                sessionStorage.setItem("last_address", data.address || "");
                            } catch (e) {}
                        }
                    } catch (e) {}
                }, 250);
            }
        });
    }

    // Informative Save-As-PDF filename dispatcher for official marksheet
    window.printMarksheet = function() {
        const instName = document.getElementById("marksheetHeaderInstName")?.textContent?.trim() || "Cumilla Educational Institution";
        const rawEiin = document.getElementById("marksheetHeaderEiin")?.textContent?.trim() || "";
        const eiin = (rawEiin && rawEiin.toUpperCase() !== "N/A" && rawEiin.toUpperCase() !== "EIIN") ? rawEiin : "";
        const studentName = document.getElementById("marksheetName")?.textContent?.trim() || "Student";
        const roll = document.getElementById("marksheetRoll")?.textContent?.trim() || "";
        const exam = document.getElementById("marksheetExam")?.textContent?.trim() || "HSC";
        const year = document.getElementById("marksheetYear")?.textContent?.trim() || "";

        const cleanInst = instName.replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
        const cleanStudent = studentName.replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
        const rollPart = roll && roll !== "-" ? ` (Roll ${roll})` : "";
        const examPart = [exam, year].filter(Boolean).join(" ").trim();
        const eiinPrefix = eiin ? `${eiin}, ` : "";

        // Standardized save file name:
        const saveTitle = `${eiinPrefix}${cleanInst} - ${cleanStudent}${rollPart} Marksheet - ${examPart}`.trim();

        const origTitle = document.title;
        document.title = saveTitle;
        window.print();
        setTimeout(() => {
            document.title = origTitle;
        }, 1500);
    };

    // --- 0. Dynamic Gateway Notice (SSC vs HSC) ---
    function updateCumillaGatewayUI() {
        const examVal = (examSelect ? examSelect.value : "hsc").toLowerCase();
        const isSsc = examVal.includes("ssc") || examVal.includes("dakhil");
        const portalUrl = isSsc ? "https://result19.comillaboard.gov.bd/SSCRCBY26/individual/" : "https://hscresult.comillaboard.gov.bd/";
        const portalDomain = isSsc ? "result19.comillaboard.gov.bd/SSCRCBY26/individual/" : "hscresult.comillaboard.gov.bd";
        const examLabel = isSsc ? "SSC" : "HSC";

        const cumillaActiveGateway = document.getElementById("cumillaActiveGateway");
        const serverNoticeSubtitle = document.getElementById("serverNoticeSubtitle");
        const indRollInput = document.getElementById("indRoll");
        const testRollsContainer = document.getElementById("testRollsContainer");
        const marksheetExam = document.getElementById("marksheetExam");

        if (cumillaActiveGateway) {
            cumillaActiveGateway.textContent = portalDomain;
            cumillaActiveGateway.href = portalUrl;
        }
        if (serverNoticeSubtitle) {
            serverNoticeSubtitle.textContent = `Live records fetched from ${portalDomain}`;
        }
        if (marksheetExam) {
            marksheetExam.textContent = examLabel;
        }
        if (indRollInput) {
            indRollInput.placeholder = "e.g. 160000";
        }
        if (testRollsContainer) {
            if (isSsc) {
                testRollsContainer.innerHTML = `
                    <span class="text-muted fs-8">Test:</span>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160000')">160000</button>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160001')">160001</button>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160002')">160002</button>
                `;
            } else {
                testRollsContainer.innerHTML = `
                    <span class="text-muted fs-8">Test:</span>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160000')">160000</button>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160001')">160001</button>
                    <button type="button" class="btn btn-outline-secondary btn-xs py-0 px-1.5 rounded" onclick="setTestRoll('160002')">160002</button>
                `;
            }
        }
    }

    if (examSelect) {
        examSelect.addEventListener("change", updateCumillaGatewayUI);
        updateCumillaGatewayUI();
    }

    // --- 1. Drag & Drop File Upload ---
    if (dropZone && fileInput) {
        ['dragenter', 'dragover'].forEach(name => {
            dropZone.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.add('dragover');
            });
        });

        ['dragleave', 'drop'].forEach(name => {
            dropZone.addEventListener(name, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.remove('dragover');
            });
        });

        dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                displaySelectedFile(files[0].name);
            }
        });

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                displaySelectedFile(fileInput.files[0].name);
            }
        });
    }

    function displaySelectedFile(name) {
        if (fileName) fileName.textContent = name;
        if (fileInfo) fileInfo.classList.remove('d-none');
    }

    // --- 2. Single Roll Instant Lookup (NO REG & NO CAPTCHA) ---
    if (cumillaSingleForm) {
        cumillaSingleForm.addEventListener("submit", async (e) => {
            e.preventDefault();

            const roll = indRollInput.value.trim();
            const exam = examSelect ? examSelect.value : "hsc";
            const year = yearSelect ? yearSelect.value : "2026";

            if (!roll) {
                alert("Please enter a valid Cumilla roll number.");
                return;
            }

            submitIndBtn.disabled = true;
            submitIndBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Fetching Result...`;

            const formData = new FormData();
            formData.append("exam", exam);
            formData.append("year", year);
            formData.append("board", board);
            formData.append("roll", roll);
            formData.append("reg", "");
            formData.append("captcha", "");

            try {
                const response = await fetch("/api/check-individual", {
                    method: "POST",
                    body: formData
                });

                if (!response.ok) {
                    const errData = await response.json();
                    throw new Error(errData.detail || "Failed to fetch student result from Cumilla server.");
                }

                const result = await response.json();
                if (result.success && result.data) {
                    renderIndividualMarksheet(result.data);
                    appendResultRow(result.data);
                } else {
                    alert("No result found for Roll " + roll + " under Cumilla Board.");
                }
            } catch (err) {
                console.error("Cumilla Individual Search Error:", err);
                alert(err.message || "Connection to Cumilla server timed out. Please try again.");
            } finally {
                submitIndBtn.disabled = false;
                submitIndBtn.innerHTML = `<i class="bi bi-search me-1"></i> Get Cumilla Marksheet`;
            }
        });
    }

    function renderIndividualMarksheet(student) {
        if (!individualResultCard) return;
        // Dynamically resolve Institutional Header Details (Name, Address, EIIN)
        const info = resolveInstitutionInfo(
            student.Institute,
            student.EIIN || student.eiin,
            student.Institution_Address || student.address || ""
        );
        updateMarksheetInstitutionalHeader(info);
        if (!info.eiin && student.Institute) {
            fetch(`/api/lookup-institution?institute=${encodeURIComponent(student.Institute)}&board=Cumilla`)
                .then(r => r.ok ? r.json() : null)
                .then(data => {
                    if (data && data.found && data.eiin) {
                        updateMarksheetInstitutionalHeader({
                            name: info.name || data.name,
                            eiin: data.eiin,
                            address: data.address || info.address
                        });
                    }
                })
                .catch(() => {});
        }

        marksheetName.textContent = student.Name || "STUDENT NAME";
        marksheetExam.textContent = (student.Exam || "HSC").toUpperCase();
        marksheetYear.textContent = student.Year || "2026";
        marksheetBoard.textContent = "CUMILLA";
        marksheetRoll.textContent = student.Roll || "-";
        marksheetReg.textContent = student.Registration || "N/A";
        marksheetGroup.textContent = student.Group || "GENERAL";
        marksheetType.textContent = student.Type || "REGULAR";
        marksheetFather.textContent = student["Father Name"] || student.Father_Name || "-";
        marksheetMother.textContent = student["Mother Name"] || student.Mother_Name || "-";
        marksheetInstitute.textContent = student.Institute || info.name || "-";

        const isPassed = (student.Result || "").toUpperCase().includes("PASSED");
        const gpaVal = student.GPA || "0.00";

        marksheetGpa.textContent = gpaVal;
        if (isPassed) {
            marksheetResultBadge.className = "badge bg-success-subtle text-success fs-7 px-3 py-2 border border-success-subtle d-inline-block";
            marksheetResultBadge.innerHTML = `RESULT: PASSED &bull; GPA: <strong id="marksheetGpa">${gpaVal}</strong>`;
        } else {
            marksheetResultBadge.className = "badge bg-danger-subtle text-danger fs-7 px-3 py-2 border border-danger-subtle d-inline-block";
            marksheetResultBadge.innerHTML = `RESULT: ${student.Result || 'FAILED'}`;
        }

        marksheetGradesBody.innerHTML = "";
        const grades = student.Subject_Grades || {};
        const entries = Object.entries(grades);

        if (entries.length > 0) {
            entries.forEach(([subName, subGrade]) => {
                const tr = document.createElement("tr");
                let badgeClass = "badge bg-light text-dark border";
                if (subGrade === "A+") badgeClass = "badge bg-success text-white";
                else if (subGrade === "A" || subGrade === "A-") badgeClass = "badge bg-success-subtle text-success border border-success-subtle";
                else if (subGrade === "F") badgeClass = "badge bg-danger text-white";

                tr.innerHTML = `
                    <td class="text-start py-1.5">${subName}</td>
                    <td class="py-1.5"><span class="${badgeClass} px-2 py-0.5 rounded">${subGrade}</span></td>
                `;
                marksheetGradesBody.appendChild(tr);
            });
        } else {
            marksheetGradesBody.innerHTML = `
                <tr><td colspan="2" class="text-center text-muted py-2">Subject details not available</td></tr>
            `;
        }

        individualResultCard.style.display = "block";
        individualResultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }

    // --- 3. Batch File Upload ---
    if (cumillaBatchFileForm) {
        cumillaBatchFileForm.addEventListener("submit", async (e) => {
            e.preventDefault();

            if (!fileInput.files || fileInput.files.length === 0) {
                alert("Please select or drop a CSV or Excel file containing student rolls.");
                return;
            }

            const exam = examSelect ? examSelect.value : "hsc";
            const year = yearSelect ? yearSelect.value : "2026";
            const isDemo = demoModeCheck ? demoModeCheck.checked : false;

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);
            formData.append("exam", exam);
            formData.append("year", year);
            formData.append("board", board);
            formData.append("demo_mode", isDemo ? "true" : "false");

            submitBulkBtn.disabled = true;
            submitBulkBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Initializing...`;

            resetExtractionUI();

            try {
                const resp = await fetch("/api/start-extraction", {
                    method: "POST",
                    body: formData
                });

                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.detail || "Failed to start Cumilla extraction.");
                }

                const data = await resp.json();
                connectSse(data.job_id, data.total_students, submitBulkBtn, `<i class="bi bi-play-circle-fill me-1"></i> Start Batch Extraction`);

            } catch (err) {
                alert(err.message || "Failed to start extraction.");
                progressCard.style.display = "none";
                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Batch Extraction`;
            }
        });
    }

    // --- 4. Paste Rolls Directly ---
    if (cumillaPasteForm) {
        cumillaPasteForm.addEventListener("submit", async (e) => {
            e.preventDefault();

            const rawText = pastedRollsInput ? pastedRollsInput.value.trim() : "";
            if (!rawText) {
                alert("Please enter or paste at least one roll number.");
                return;
            }

            const exam = examSelect ? examSelect.value : "hsc";
            const year = yearSelect ? yearSelect.value : "2026";
            const isDemo = pasteDemoModeCheck ? pasteDemoModeCheck.checked : false;

            const formData = new FormData();
            formData.append("rolls", rawText);
            formData.append("exam", exam);
            formData.append("year", year);
            formData.append("board", board);
            formData.append("demo_mode", isDemo ? "true" : "false");

            submitPasteBtn.disabled = true;
            submitPasteBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Processing Rolls...`;

            resetExtractionUI();

            try {
                const resp = await fetch("/api/start-extraction-rolls", {
                    method: "POST",
                    body: formData
                });

                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.detail || "Failed to process pasted rolls.");
                }

                const data = await resp.json();
                connectSse(data.job_id, data.total_students, submitPasteBtn, `<i class="bi bi-cpu me-1"></i> Extract Pasted Rolls`);

            } catch (err) {
                alert(err.message || "Error processing pasted rolls.");
                progressCard.style.display = "none";
                submitPasteBtn.disabled = false;
                submitPasteBtn.innerHTML = `<i class="bi bi-cpu me-1"></i> Extract Pasted Rolls`;
            }
        });
    }

    function resetExtractionUI() {
        extractedCount = 0;
        resultsBody.innerHTML = "";
        resultsCount.textContent = "0 Students";
        progressCard.style.display = "block";
        progressTitle.textContent = "Extracting Cumilla results via direct portal...";
        progressBar.style.width = "0%";
        progressBar.className = "progress-bar progress-bar-striped progress-bar-animated bg-success";
        progressPercentBadge.textContent = "0%";
        progressPercentBadge.className = "badge bg-success rounded-pill px-2.5 py-1";
        statusMessage.innerHTML = "Initializing fast connection...";
        studentCounter.textContent = "0 / 0";
        downloadSection.style.setProperty("display", "none", "important");
    }

    function connectSse(jobId, totalStudents, triggerBtn, originalBtnHtml) {
        if (eventSource) eventSource.close();

        eventSource = new EventSource(`/api/progress-stream/${jobId}`);

        eventSource.onmessage = (e) => {
            let msg;
            try {
                msg = JSON.parse(e.data);
            } catch (err) {
                return;
            }

            if (msg.type === "progress") {
                const cur = msg.current;
                const tot = msg.total || totalStudents;
                const pct = msg.percent;
                const res = msg.data;

                progressBar.style.width = `${pct}%`;
                progressPercentBadge.textContent = `${pct}%`;
                studentCounter.textContent = `${cur} / ${tot}`;
                statusMessage.innerHTML = `<i class="bi bi-lightning-charge text-success me-1"></i> Scraped: Roll <strong>${res.Roll}</strong> (${res.Name})`;

                appendResultRow(res);

            } else if (msg.type === "completed") {
                progressBar.style.width = "100%";
                progressBar.className = "progress-bar bg-success";
                progressPercentBadge.textContent = "100%";
                progressPercentBadge.className = "badge bg-success rounded-pill px-2.5 py-1";
                progressTitle.textContent = "Cumilla Results Extracted Successfully!";
                statusMessage.innerHTML = `<i class="bi bi-check-circle-fill text-success me-1"></i> Loaded ${msg.total} student records via direct portal!`;

                downloadXlsxBtn.href = msg.download_xlsx;
                downloadCsvBtn.href = msg.download_csv;
                if (msg.excel_filename) downloadXlsxBtn.setAttribute("download", msg.excel_filename);
                if (msg.csv_filename) downloadCsvBtn.setAttribute("download", msg.csv_filename);
                downloadSection.style.setProperty("display", "flex", "important");

                sessionStorage.setItem("last_job_id", jobId);
                if (msg.analysis) {
                    sessionStorage.setItem("last_analysis", JSON.stringify(msg.analysis));
                    try {
                        if (msg.analysis.eiin) sessionStorage.setItem("last_eiin", msg.analysis.eiin);
                        else sessionStorage.removeItem("last_eiin");
                        sessionStorage.setItem("last_inst", msg.analysis.institute || "");
                        sessionStorage.setItem("last_address", msg.analysis.address || "");
                    } catch (e) {}
                }
                if (openAnalysisBtn) {
                    openAnalysisBtn.href = `/result-analysis?job_id=${jobId}`;
                }

                triggerBtn.disabled = false;
                triggerBtn.innerHTML = originalBtnHtml;
                eventSource.close();

            } else if (msg.type === "error") {
                progressBar.className = "progress-bar bg-danger";
                statusMessage.innerHTML = `<i class="bi bi-exclamation-triangle-fill text-danger me-1"></i> ${msg.message}`;
                alert(msg.message);

                triggerBtn.disabled = false;
                triggerBtn.innerHTML = originalBtnHtml;
                eventSource.close();
            }
        };

        eventSource.onerror = () => {
            console.error("SSE stream error");
        };
    }

    function appendResultRow(data) {
        const emptyRow = document.getElementById("emptyResultsRow");
        if (emptyRow) emptyRow.remove();

        extractedCount++;
        resultsCount.textContent = `${extractedCount} Students`;

        const row = document.createElement("tr");
        const isPassed = (data.Result || "").toUpperCase().includes("PASSED");
        const gpa = data.GPA || "0.00";

        let gpaBadge = `<span class="badge bg-secondary-subtle text-secondary">${gpa}</span>`;
        if (parseFloat(gpa) === 5.0) {
            gpaBadge = `<span class="badge bg-warning text-dark fw-bold">5.00 A+</span>`;
        } else if (parseFloat(gpa) >= 4.0) {
            gpaBadge = `<span class="badge bg-success text-white fw-semibold">${gpa}</span>`;
        } else if (parseFloat(gpa) > 0) {
            gpaBadge = `<span class="badge bg-info-subtle text-dark">${gpa}</span>`;
        }

        const resultBadge = isPassed
            ? `<span class="badge bg-success-subtle text-success border border-success-subtle"><i class="bi bi-check-circle me-1"></i>PASSED</span>`
            : `<span class="badge bg-danger-subtle text-danger border border-danger-subtle"><i class="bi bi-x-circle me-1"></i>${data.Result || 'FAILED'}</span>`;

        row.innerHTML = `
            <td><strong>${data.Roll || '-'}</strong></td>
            <td><strong class="text-dark">${data.Name || 'N/A'}</strong></td>
            <td><small class="text-muted">${data.Group || 'GENERAL'}</small></td>
            <td>${gpaBadge}</td>
            <td>${resultBadge}</td>
            <td><small class="text-muted text-truncate d-inline-block" style="max-width: 220px;" title="${data.Subject_Grades_Summary || ''}">${data.Subject_Grades_Summary || '-'}</small></td>
        `;

        resultsBody.prepend(row);
    }
});
