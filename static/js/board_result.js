/**
 * board_result.js - Dedicated Script for General Education Board Results
 * Handles captcha fetching, single roll verification, batch file upload with SSE live streaming.
 */

document.addEventListener("DOMContentLoaded", () => {
    // Examination inputs
    const examSelect = document.getElementById("exam");
    const yearSelect = document.getElementById("year");
    const boardSelect = document.getElementById("board");
    const activeBoardName = document.getElementById("activeBoardName");

    // Single student check elements
    const individualForm = document.getElementById("individualForm");
    const indRollInput = document.getElementById("indRoll");
    const indRegInput = document.getElementById("indReg");
    const indCaptchaInput = document.getElementById("indCaptcha");
    const captchaContainer = document.getElementById("captchaContainer");
    const captchaImg = document.getElementById("captchaImg");
    const captchaSpinner = document.getElementById("captchaSpinner");
    const refreshCaptchaBtn = document.getElementById("refreshCaptchaBtn");
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

    // Batch upload elements
    const bulkForm = document.getElementById("bulkForm");
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const fileInfo = document.getElementById("fileInfo");
    const fileName = document.getElementById("fileName");
    const demoModeCheck = document.getElementById("demoMode");
    const submitBulkBtn = document.getElementById("submitBulkBtn");

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
    let lastCaptchaLoadTime = 0;

    // --- 1. Captcha Loader ---
    async function loadCaptcha(force = false) {
        const now = Date.now();
        if (!force && now - lastCaptchaLoadTime < 4000) return;
        lastCaptchaLoadTime = now;

        if (captchaImg) captchaImg.style.display = "none";
        if (captchaSpinner) captchaSpinner.style.display = "inline-block";

        try {
            const resp = await fetch("/api/captcha");
            const data = await resp.json();
            if (data.success && data.image) {
                if (captchaImg) {
                    captchaImg.src = data.image;
                    captchaImg.style.display = "inline-block";
                }
                if (data.predicted && indCaptchaInput && !indCaptchaInput.value) {
                    indCaptchaInput.value = data.predicted;
                }
            }
        } catch (err) {
            console.warn("Could not load security code:", err);
        } finally {
            if (captchaSpinner) captchaSpinner.style.display = "none";
        }
    }

    if (refreshCaptchaBtn) {
        refreshCaptchaBtn.addEventListener("click", () => loadCaptcha(true));
    }

    // Update board name badge and reload captcha on board change
    if (boardSelect) {
        boardSelect.addEventListener("change", () => {
            if (activeBoardName) {
                const text = boardSelect.options[boardSelect.selectedIndex]?.text || "Dhaka";
                activeBoardName.textContent = text.replace(/\(.*\)/, "").trim();
            }
            loadCaptcha(true);
        });
    }

    // Initial captcha load
    loadCaptcha();

    // --- 2. Drag & Drop File Upload ---
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

    // --- 3. Single Student Result Lookup ---
    if (individualForm) {
        individualForm.addEventListener("submit", async (e) => {
            e.preventDefault();

            const roll = indRollInput.value.trim();
            const reg = indRegInput.value.trim();
            const captcha = indCaptchaInput ? indCaptchaInput.value.trim() : "";
            const exam = examSelect.value;
            const year = yearSelect.value;
            const board = boardSelect.value;

            if (!roll) {
                alert("Please enter a valid Roll Number.");
                return;
            }

            submitIndBtn.disabled = true;
            submitIndBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Searching...`;

            const formData = new FormData();
            formData.append("exam", exam);
            formData.append("year", year);
            formData.append("board", board);
            formData.append("roll", roll);
            formData.append("reg", reg);
            formData.append("captcha", captcha);
            if (demoModeCheck && demoModeCheck.checked) {
                formData.append("demo_mode", "true");
            }

            try {
                const response = await fetch("/api/check-individual", {
                    method: "POST",
                    body: formData
                });

                if (!response.ok) {
                    const errData = await response.json();
                    throw new Error(errData.detail || "Failed to fetch student result.");
                }

                const result = await response.json();
                if (result.success && result.data) {
                    renderIndividualMarksheet(result.data);
                    appendResultRow(result.data);
                } else {
                    alert("No result found for the provided examination credentials.");
                }
            } catch (err) {
                console.error("Individual Search Error:", err);
                alert(err.message || "Could not retrieve result. Please verify credentials or try again.");
                loadCaptcha(true);
            } finally {
                submitIndBtn.disabled = false;
                submitIndBtn.innerHTML = `<i class="bi bi-search me-1"></i> Get Official Marksheet`;
            }
        });
    }

    function renderIndividualMarksheet(student) {
        if (!individualResultCard) return;
        // Populate Institutional Header Details (Name, Address, EIIN)
        const rawInst = student.Institute || "";
        let instName = rawInst;
        let eiinVal = "105820";
        if (rawInst) {
            const m = rawInst.match(/\(?\b(\d{5,7})\b\)?/);
            if (m) {
                eiinVal = m[1];
                instName = rawInst.replace(/\(?\b\d{5,7}\b\)?/g, "").replace(/[-()]/g, "").trim();
            }
        }
        const marksheetHeaderInstName = document.getElementById("marksheetHeaderInstName");
        const marksheetHeaderAddress = document.getElementById("marksheetHeaderAddress");
        const marksheetHeaderEiin = document.getElementById("marksheetHeaderEiin");
        if (marksheetHeaderInstName) marksheetHeaderInstName.textContent = instName || "EDUCATIONAL INSTITUTION";
        if (marksheetHeaderEiin) marksheetHeaderEiin.textContent = eiinVal;
        try { sessionStorage.setItem("last_eiin", eiinVal); } catch (e) {}
        if (marksheetHeaderAddress) {
            const b = (student.Board || "Dhaka").toLowerCase();
            marksheetHeaderAddress.textContent = b.includes("cumilla") || b.includes("comilla") ? "Kotbari Road, Cumilla Sadar, Cumilla, Bangladesh" : (b.includes("dhaka") ? "Dhaka - 1205, Bangladesh" : `${student.Board || "Education Board"}, Bangladesh`);
        }


        marksheetName.textContent = student.Name || "STUDENT NAME";
        marksheetExam.textContent = (student.Exam || "HSC").toUpperCase();
        marksheetYear.textContent = student.Year || "2026";
        marksheetBoard.textContent = (student.Board || "Dhaka").toUpperCase();
        marksheetRoll.textContent = student.Roll || "-";
        marksheetReg.textContent = student.Registration || "-";
        marksheetGroup.textContent = student.Group || "GENERAL";
        marksheetType.textContent = student.Type || "REGULAR";
        marksheetFather.textContent = student["Father Name"] || student.Father_Name || "-";
        marksheetMother.textContent = student["Mother Name"] || student.Mother_Name || "-";
        marksheetInstitute.textContent = student.Institute || "-";

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

        // Render Subject Grades
        marksheetGradesBody.innerHTML = "";
        const grades = student.Subject_Grades || {};
        const entries = Object.entries(grades);

        if (entries.length > 0) {
            entries.forEach(([subName, subGrade]) => {
                const tr = document.createElement("tr");
                let badgeClass = "badge bg-light text-dark border";
                if (subGrade === "A+") badgeClass = "badge bg-primary text-white";
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

    // --- 4. Batch File Upload & SSE Live Streaming ---
    if (bulkForm) {
        bulkForm.addEventListener("submit", async (e) => {
            e.preventDefault();

            if (!fileInput.files || fileInput.files.length === 0) {
                alert("Please select or drop a CSV or Excel file first.");
                return;
            }

            const exam = examSelect.value;
            const year = yearSelect.value;
            const board = boardSelect.value;
            const isDemo = demoModeCheck ? demoModeCheck.checked : false;

            const formData = new FormData();
            formData.append("file", fileInput.files[0]);
            formData.append("exam", exam);
            formData.append("year", year);
            formData.append("board", board);
            formData.append("demo_mode", isDemo ? "true" : "false");

            submitBulkBtn.disabled = true;
            submitBulkBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Initializing Extraction...`;

            // Reset UI
            extractedCount = 0;
            resultsBody.innerHTML = "";
            resultsCount.textContent = "0 Students";
            progressCard.style.display = "block";
            progressTitle.textContent = "Processing student results...";
            progressBar.style.width = "0%";
            progressBar.className = "progress-bar progress-bar-striped progress-bar-animated bg-primary";
            progressPercentBadge.textContent = "0%";
            progressPercentBadge.className = "badge bg-primary rounded-pill px-2.5 py-1";
            statusMessage.innerHTML = "Initializing extraction job...";
            studentCounter.textContent = "0 / 0";
            downloadSection.style.setProperty("display", "none", "important");

            try {
                const resp = await fetch("/api/start-extraction", {
                    method: "POST",
                    body: formData
                });

                if (!resp.ok) {
                    const err = await resp.json();
                    throw new Error(err.detail || "Failed to initialize extraction.");
                }

                const data = await resp.json();
                const jobId = data.job_id;
                const total = data.total_students;
                studentCounter.textContent = `0 / ${total}`;

                // Connect SSE
                connectSse(jobId, total);

            } catch (err) {
                alert(err.message || "An error occurred starting extraction.");
                progressCard.style.display = "none";
                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Batch Extraction`;
            }
        });
    }

    function connectSse(jobId, totalStudents) {
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
                statusMessage.innerHTML = `<i class="bi bi-gear spin me-1"></i> Processed: Roll <strong>${res.Roll}</strong> (${res.Name})`;

                appendResultRow(res);

            } else if (msg.type === "completed") {
                progressBar.style.width = "100%";
                progressBar.className = "progress-bar bg-success";
                progressPercentBadge.textContent = "100%";
                progressPercentBadge.className = "badge bg-success rounded-pill px-2.5 py-1";
                progressTitle.textContent = "Extraction Completed Successfully!";
                statusMessage.innerHTML = `<i class="bi bi-check-circle-fill text-success me-1"></i> Successfully loaded ${msg.total} student records!`;

                // Set Download Links
                downloadXlsxBtn.href = msg.download_xlsx;
                downloadCsvBtn.href = msg.download_csv;
                downloadSection.style.setProperty("display", "flex", "important");

                // Store analysis data for Result Analysis page
                sessionStorage.setItem("last_job_id", jobId);
                if (msg.analysis) {
                    sessionStorage.setItem("last_analysis", JSON.stringify(msg.analysis));
                    if (msg.analysis.eiin) {
                        try { sessionStorage.setItem("last_eiin", msg.analysis.eiin); } catch (e) {}
                    }
                }
                if (openAnalysisBtn) {
                    openAnalysisBtn.href = `/result-analysis?job_id=${jobId}`;
                }

                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Batch Extraction`;
                eventSource.close();

            } else if (msg.type === "error") {
                progressBar.className = "progress-bar bg-danger";
                statusMessage.innerHTML = `<i class="bi bi-exclamation-triangle-fill text-danger me-1"></i> ${msg.message}`;
                alert(msg.message);

                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Batch Extraction`;
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
            gpaBadge = `<span class="badge bg-primary text-white fw-semibold">${gpa}</span>`;
        } else if (parseFloat(gpa) > 0) {
            gpaBadge = `<span class="badge bg-info-subtle text-dark">${gpa}</span>`;
        }

        const resultBadge = isPassed
            ? `<span class="badge bg-success-subtle text-success border border-success-subtle"><i class="bi bi-check-circle me-1"></i>PASSED</span>`
            : `<span class="badge bg-danger-subtle text-danger border border-danger-subtle"><i class="bi bi-x-circle me-1"></i>${data.Result || 'FAILED'}</span>`;

        row.innerHTML = `
            <td><strong>${data.Roll || '-'}</strong></td>
            <td><small class="text-muted">${data.Registration || '-'}</small></td>
            <td><strong class="text-dark">${data.Name || 'N/A'}</strong></td>
            <td>${gpaBadge}</td>
            <td>${resultBadge}</td>
            <td><small class="text-muted text-truncate d-inline-block" style="max-width: 220px;" title="${data.Subject_Grades_Summary || ''}">${data.Subject_Grades_Summary || '-'}</small></td>
        `;

        resultsBody.prepend(row);
    }
});
