// main.js - Frontend Logic for Board Result Analyzer (Bulk, Individual & Comprehensive Analysis)

document.addEventListener("DOMContentLoaded", () => {
    // Shared Controls
    const examSelect = document.getElementById("exam");
    const yearSelect = document.getElementById("year");
    const boardSelect = document.getElementById("board");
    const demoModeCheck = document.getElementById("demoMode");

    // Bulk Mode Elements
    const bulkForm = document.getElementById("bulkForm");
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const fileInfo = document.getElementById("fileInfo");
    const fileName = document.getElementById("fileName");
    const submitBulkBtn = document.getElementById("submitBulkBtn");

    // Individual Mode Elements
    const individualForm = document.getElementById("individualForm");
    const indRollInput = document.getElementById("indRoll");
    const indRegInput = document.getElementById("indReg");
    const submitIndBtn = document.getElementById("submitIndBtn");
    const captchaContainer = document.getElementById("captchaContainer");
    const captchaImg = document.getElementById("captchaImg");
    const captchaSpinner = document.getElementById("captchaSpinner");
    const refreshCaptchaBtn = document.getElementById("refreshCaptchaBtn");
    const indCaptchaInput = document.getElementById("indCaptcha");

    // Marksheet Card Elements
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
    const marksheetInstituteRow = document.getElementById("marksheetInstituteRow");
    const marksheetGradesBody = document.getElementById("marksheetGradesBody");
    const cumillaHscAlert = document.getElementById("cumillaHscAlert");
    const eduBoardAlert = document.getElementById("eduBoardAlert");
    const activeBoardName = document.getElementById("activeBoardName");
    const regHelpText = document.getElementById("regHelpText");

    // Progress Elements (Bulk Scraping)
    const progressCard = document.getElementById("progressCard");
    const progressTitle = document.getElementById("progressTitle");
    const progressBar = document.getElementById("progressBar");
    const progressPercentBadge = document.getElementById("progressPercentBadge");
    const statusMessage = document.getElementById("statusMessage");
    const studentCounter = document.getElementById("studentCounter");

    // Download Elements
    const downloadSection = document.getElementById("downloadSection");
    const downloadXlsxBtn = document.getElementById("downloadXlsxBtn");
    const downloadCsvBtn = document.getElementById("downloadCsvBtn");

    // Results Table Elements
    const resultsBody = document.getElementById("resultsBody");
    const resultsCount = document.getElementById("resultsCount");

    // Result Analysis Buttons & Elements
    const openAnalysisBtn = document.getElementById("openAnalysisBtn");
    const headerAnalysisBtn = document.getElementById("headerAnalysisBtn");
    const bottomAnalysisBtn = document.getElementById("bottomAnalysisBtn");
    const uploadFileForAnalysisBtn = document.getElementById("uploadFileForAnalysisBtn");
    const analysisFileInput = document.getElementById("analysisFileInput");
    const exportAnalysisExcelBtn = document.getElementById("exportAnalysisExcelBtn");
    const printAnalysisBtn = document.getElementById("printAnalysisBtn");

    // In-memory data store for live analysis
    let eventSource = null;
    let extractedCount = 0;
    let currentRecords = [];
    let latestAnalysisData = null;
    let latestJobId = null;
    let analysisModalInstance = null;

    // Bootstrap Modal initialization
    const analysisModalEl = document.getElementById("resultAnalysisModal");
    if (analysisModalEl && typeof bootstrap !== "undefined") {
        analysisModalInstance = new bootstrap.Modal(analysisModalEl);
    }

    if (printAnalysisBtn) {
        printAnalysisBtn.onclick = () => {
            if (!latestAnalysisData) {
                alert("Please upload a result file or run analysis first before generating the PDF report.");
                return;
            }
            printAnalysisReport(latestAnalysisData, {
                exam: (examSelect && examSelect.options[examSelect.selectedIndex]?.text) || "HSC",
                year: (yearSelect && yearSelect.value) || "2026",
                board: (boardSelect && boardSelect.options[boardSelect.selectedIndex]?.text) || "Board"
            });
        };
    }

    let lastCaptchaLoadTime = 0;

    async function loadCaptcha(force = false) {
        if (!captchaContainer || captchaContainer.style.display === "none") return;
        const now = Date.now();
        if (!force && now - lastCaptchaLoadTime < 3000) return;
        lastCaptchaLoadTime = now;

        if (captchaSpinner) captchaSpinner.style.display = "inline-block";
        if (captchaImg) captchaImg.style.display = "none";
        try {
            const resp = await fetch("/api/captcha");
            const data = await resp.json();
            if (data.success && data.image) {
                if (captchaImg) {
                    captchaImg.src = data.image;
                    captchaImg.style.display = "inline-block";
                }
                if (captchaSpinner) captchaSpinner.style.display = "none";
                if (indCaptchaInput) {
                    indCaptchaInput.value = data.predicted || "";
                }
            } else {
                if (captchaSpinner) captchaSpinner.style.display = "none";
            }
        } catch (e) {
            console.warn("Failed to load captcha:", e);
            if (captchaSpinner) captchaSpinner.style.display = "none";
        }
    }

    if (refreshCaptchaBtn) {
        refreshCaptchaBtn.addEventListener("click", () => {
            loadCaptcha(true);
        });
    }

    // Dynamic Server Notice for Cumilla HSC/SSC vs National Portal
    function updateCumillaServerNotice() {
        const boardVal = (boardSelect.value || "").toLowerCase();
        const isCumilla = boardVal === "comilla" || boardVal === "cumilla";
        const examVal = (examSelect.value || "").toLowerCase();
        const isHsc = examVal.includes("hsc");
        const isSsc = examVal.includes("ssc") || examVal.includes("dakhil");
        const selectedBoardText = boardSelect.options[boardSelect.selectedIndex]?.text || "Selected Board";

        if (isCumilla && (isHsc || isSsc)) {
            const portalUrl = isSsc ? "https://result19.comillaboard.gov.bd/SSCRCBY26/individual/" : "https://hscresult.comillaboard.gov.bd/";
            const portalDomain = isSsc ? "result19.comillaboard.gov.bd/SSCRCBY26/individual/" : "hscresult.comillaboard.gov.bd";
            const examLabel = isSsc ? "SSC" : "HSC";

            if (cumillaHscAlert) {
                cumillaHscAlert.style.display = "flex";
                const bannerText = document.getElementById("cumillaBannerText");
                if (bannerText) {
                    bannerText.innerHTML = `Connected to <a href="${portalUrl}" target="_blank" class="fw-semibold text-success text-decoration-underline">${portalDomain}</a>. Only <strong>Roll number</strong> is needed for Cumilla ${examLabel}. Registration number is not required.`;
                }
            }
            if (eduBoardAlert) eduBoardAlert.style.display = "none";
            if (captchaContainer) captchaContainer.style.display = "none";
            indRollInput.placeholder = isSsc ? "e.g. 100000" : "e.g. 160000";
            indRegInput.placeholder = `Optional (Not needed for Cumilla ${examLabel})`;
            if (regHelpText) regHelpText.textContent = `Optional. For Cumilla ${examLabel}, result is found automatically with Roll only.`;
        } else {
            if (cumillaHscAlert) cumillaHscAlert.style.display = "none";
            if (eduBoardAlert) {
                eduBoardAlert.style.display = "flex";
                if (activeBoardName) {
                    activeBoardName.textContent = selectedBoardText.replace(/\(.*\)/, "").trim();
                }
            }
            if (captchaContainer) {
                captchaContainer.style.display = "block";
                loadCaptcha();
            }
            indRollInput.placeholder = "e.g. 100101";
            indRegInput.placeholder = "e.g. 1512345678 (Registration No)";
            if (regHelpText) regHelpText.textContent = "Registration is required for full subject-wise marksheet on national portal.";
        }
    }

    examSelect.addEventListener("change", updateCumillaServerNotice);
    boardSelect.addEventListener("change", updateCumillaServerNotice);
    updateCumillaServerNotice();

    // --- Drag and Drop Handlers for Bulk File ---
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
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

    function displaySelectedFile(name) {
        fileName.textContent = name;
        fileInfo.classList.remove('d-none');
    }

    // ==========================================
    // 1. INDIVIDUAL STUDENT RESULT CHECK
    // ==========================================
    individualForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        const rollVal = indRollInput.value.trim();
        const regVal = indRegInput.value.trim();

        if (!rollVal) {
            alert("Please enter a valid Roll Number.");
            return;
        }

        // UI Loading State
        submitIndBtn.disabled = true;
        submitIndBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status"></span>Fetching Result...`;
        individualResultCard.style.display = "none";

        const formData = new FormData();
        formData.append("exam", examSelect.value);
        formData.append("year", yearSelect.value);
        formData.append("board", boardSelect.value);
        formData.append("roll", rollVal);
        formData.append("reg", regVal);
        const captchaVal = indCaptchaInput ? indCaptchaInput.value.trim() : "";
        formData.append("captcha", captchaVal);
        formData.append("demo_mode", demoModeCheck && demoModeCheck.checked ? "true" : "false");

        try {
            const response = await fetch("/api/check-individual", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Failed to fetch student result.");
            }

            const resJson = await response.json();
            const student = resJson.data;

            if (student.Result === "Invalid Security Code" || student.Result === "Captcha Required") {
                loadCaptcha(true);
                if (indCaptchaInput) {
                    indCaptchaInput.focus();
                    indCaptchaInput.select();
                }
            }

            // Save in currentRecords
            currentRecords.push(student);
            latestAnalysisData = null; // Invalidate to recompute on next click

            // Render detailed Marksheet Card
            renderIndividualMarksheet(student);

            // Also append to the preview log table
            appendResultRow(student);

        } catch (error) {
            alert(`Error: ${error.message}`);
        } finally {
            submitIndBtn.disabled = false;
            submitIndBtn.innerHTML = `<i class="bi bi-search me-1"></i> Check Single Result`;
        }
    });

    function renderIndividualMarksheet(student) {
        marksheetName.textContent = student.Name || "N/A";
        marksheetExam.textContent = student.Exam || examSelect.options[examSelect.selectedIndex].text;
        marksheetYear.textContent = student.Year || yearSelect.value;
        marksheetBoard.textContent = student.Board || boardSelect.options[boardSelect.selectedIndex].text;
        marksheetGpa.textContent = student.GPA || "0.00";
        marksheetRoll.textContent = student.Roll || "-";
        marksheetReg.textContent = student.Registration || "N/A";
        marksheetGroup.textContent = student.Group || "-";
        marksheetType.textContent = student.Type || "-";
        marksheetFather.textContent = student["Father Name"] || "-";
        marksheetMother.textContent = student["Mother Name"] || "-";

        if (marksheetInstitute && marksheetInstituteRow) {
            if (student.Institute) {
                marksheetInstitute.textContent = student.Institute;
                marksheetInstituteRow.style.display = "block";
            } else {
                marksheetInstituteRow.style.display = "none";
            }
        }

        // Result badge color
        const isNotPublished = (student.Result || "").includes("Not Published");
        const isError = (student.Result || "").includes("Fail") || (student.Result || "").includes("Not Found") || (student.Result || "").includes("Error");

        marksheetResultBadge.textContent = student.Result || "N/A";
        if (isNotPublished) {
            marksheetResultBadge.className = "badge bg-warning-subtle text-dark fs-7 px-3 py-2 border border-warning-subtle d-inline-block mb-1";
        } else if (isError) {
            marksheetResultBadge.className = "badge bg-danger-subtle text-danger fs-7 px-3 py-2 border border-danger-subtle d-inline-block mb-1";
        } else {
            marksheetResultBadge.className = "badge bg-success-subtle text-success fs-7 px-3 py-2 border border-success-subtle d-inline-block mb-1";
        }

        // Populate Subject Grades Table
        marksheetGradesBody.innerHTML = "";
        const grades = student.Subject_Grades || {};
        const subjectKeys = Object.keys(grades);

        if (subjectKeys.length === 0) {
            marksheetGradesBody.innerHTML = `
                <tr>
                    <td colspan="2" class="text-center text-muted py-3">
                        ${student.Subject_Grades_Summary || "No subject-wise breakdown available."}
                    </td>
                </tr>
            `;
        } else {
            subjectKeys.forEach(subj => {
                const grade = grades[subj];
                let badgeClass = "badge bg-secondary";
                if (grade === "A+") badgeClass = "badge-gpa-gold";
                else if (grade === "A" || grade === "A-") badgeClass = "badge-gpa-green";
                else if (grade === "B" || grade === "C" || grade === "D") badgeClass = "badge-gpa-blue";
                else if (grade === "F") badgeClass = "badge-gpa-red";

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-medium text-dark">${subj}</td>
                    <td class="text-center"><span class="badge ${badgeClass} px-3 py-1">${grade}</span></td>
                `;
                marksheetGradesBody.appendChild(tr);
            });
        }

        // Show card and scroll into view smoothly
        individualResultCard.style.display = "block";
        individualResultCard.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }

    // ==========================================
    // 2. BULK EXTRACTION (CSV / EXCEL UPLOAD)
    // ==========================================
    bulkForm.addEventListener("submit", async (e) => {
        e.preventDefault();

        if (!fileInput.files || fileInput.files.length === 0) {
            alert("Please select or drop a CSV or Excel file containing Roll numbers.");
            return;
        }

        // Reset UI State
        extractedCount = 0;
        currentRecords = [];
        latestAnalysisData = null;
        latestJobId = null;

        resultsCount.textContent = "0 Extracted";
        resultsBody.innerHTML = "";
        downloadSection.style.setProperty("display", "none", "important");
        downloadXlsxBtn.classList.remove("pulse-success");
        individualResultCard.style.display = "none";

        submitBulkBtn.disabled = true;
        submitBulkBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status"></span>Processing Results...`;

        progressCard.style.display = "block";
        progressTitle.textContent = "Processing student results...";
        progressBar.style.width = "0%";
        progressBar.className = "progress-bar progress-bar-striped progress-bar-animated bg-primary";
        progressPercentBadge.textContent = "0%";
        statusMessage.innerHTML = `<i class="bi bi-arrow-repeat spin me-1"></i> Reading uploaded student file...`;
        studentCounter.textContent = "0 / 0";

        const formData = new FormData();
        formData.append("file", fileInput.files[0]);
        formData.append("exam", examSelect.value);
        formData.append("year", yearSelect.value);
        formData.append("board", boardSelect.value);
        formData.append("demo_mode", demoModeCheck && demoModeCheck.checked ? "true" : "false");

        try {
            const response = await fetch("/api/start-extraction", {
                method: "POST",
                body: formData
            });

            if (!response.ok) {
                const errData = await response.json();
                throw new Error(errData.detail || "Failed to start extraction task.");
            }

            const data = await response.json();
            const jobId = data.job_id;
            const totalStudents = data.total_students;
            latestJobId = jobId;

            studentCounter.textContent = `0 / ${totalStudents}`;
            statusMessage.innerHTML = `<i class="bi bi-arrow-repeat spin me-1"></i> Connecting to board result portal...`;

            // Connect to Server-Sent Events (SSE) Stream
            connectProgressStream(jobId, totalStudents);

        } catch (error) {
            alert(`Error: ${error.message}`);
            submitBulkBtn.disabled = false;
            submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Start Bulk Extraction`;
            progressCard.style.display = "none";
        }
    });

    function connectProgressStream(jobId, total) {
        if (eventSource) {
            eventSource.close();
        }

        eventSource = new EventSource(`/api/progress-stream/${jobId}`);

        eventSource.onmessage = (event) => {
            const msg = JSON.parse(event.data);

            if (msg.type === "progress") {
                const cur = msg.current;
                const tot = msg.total;
                const pct = msg.percent;
                const res = msg.data;

                // Store in memory
                currentRecords.push(res);

                // Update Progress bar
                progressBar.style.width = `${pct}%`;
                progressPercentBadge.textContent = `${pct}%`;
                studentCounter.textContent = `${cur} / ${tot}`;
                statusMessage.innerHTML = `<i class="bi bi-gear spin me-1"></i> Processed Roll: <strong>${res.Roll}</strong> (${res.Name})`;

                // Append student row to live results table
                appendResultRow(res);

            } else if (msg.type === "completed") {
                progressBar.style.width = "100%";
                progressBar.className = "progress-bar bg-success";
                progressPercentBadge.textContent = "100%";
                progressPercentBadge.className = "badge bg-success rounded-pill px-3 py-2";
                progressTitle.textContent = "Results Loaded Successfully!";
                statusMessage.innerHTML = `<i class="bi bi-check-circle-fill text-success me-1"></i> Successfully loaded ${msg.total} student records!`;

                // Set Download Links
                downloadXlsxBtn.href = msg.download_xlsx;
                downloadCsvBtn.href = msg.download_csv;
                downloadSection.style.setProperty("display", "flex", "important");
                downloadXlsxBtn.classList.add("pulse-success");

                // Capture analysis from backend if returned
                if (msg.analysis) {
                    latestAnalysisData = msg.analysis;
                }
                if (msg.download_analysis_xlsx && exportAnalysisExcelBtn) {
                    exportAnalysisExcelBtn.href = msg.download_analysis_xlsx;
                }

                // Reset Submit Button
                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Process Batch Results`;

                eventSource.close();

            } else if (msg.type === "error") {
                progressBar.className = "progress-bar bg-danger";
                statusMessage.innerHTML = `<i class="bi bi-exclamation-triangle-fill text-danger me-1"></i> ${msg.message}`;
                alert(msg.message);

                submitBulkBtn.disabled = false;
                submitBulkBtn.innerHTML = `<i class="bi bi-play-circle-fill me-1"></i> Process Batch Results`;

                eventSource.close();
            }
        };

        eventSource.onerror = (err) => {
            console.error("SSE connection error:", err);
        };
    }

    // ==========================================
    // 3. TABLE LOG HELPER
    // ==========================================
    function appendResultRow(data) {
        // Remove empty placeholder row if present
        const emptyRow = document.getElementById("emptyRow");
        if (emptyRow) {
            emptyRow.remove();
        }

        extractedCount++;
        resultsCount.textContent = `${extractedCount} Recorded`;

        const row = document.createElement("tr");

        // Format GPA Badge
        let gpaBadgeClass = "badge-gpa-blue";
        const gpaNum = parseFloat(data.GPA);
        if (gpaNum >= 5.0) {
            gpaBadgeClass = "badge-gpa-gold";
        } else if (gpaNum >= 4.0) {
            gpaBadgeClass = "badge-gpa-green";
        } else if (gpaNum === 0 || isNaN(gpaNum) || (data.Result && (data.Result.includes("Fail") || data.Result.includes("Not Found")))) {
            gpaBadgeClass = "badge-gpa-red";
        }

        // Format Result Status Badge
        let resultBadge = `<span class="badge bg-success-subtle text-success border border-success-subtle">${data.Result || "PASSED"}</span>`;
        if (data.Result && (data.Result.includes("Fail") || data.Result.includes("Not Found"))) {
            resultBadge = `<span class="badge bg-danger-subtle text-danger border border-danger-subtle">${data.Result}</span>`;
        }

        // Truncate long subject grades for table display
        const gradesSummary = data.Subject_Grades_Summary || "-";
        const displayGrades = gradesSummary.length > 50 
            ? `${gradesSummary.substring(0, 48)}...` 
            : gradesSummary;

        const studentBoard = data.Board || (boardSelect.options[boardSelect.selectedIndex]?.text.replace(/\(.*\)/, "").trim()) || "-";

        row.innerHTML = `
            <td class="ps-4 fw-bold text-dark">${data.Roll || "-"}</td>
            <td class="text-secondary">${data.Registration || "-"}</td>
            <td class="fw-medium">${data.Name || "N/A"}</td>
            <td><span class="badge bg-light text-dark border">${studentBoard}</span></td>
            <td><span class="badge ${gpaBadgeClass} px-2 py-1 fs-7">${data.GPA || "0.00"}</span></td>
            <td>${resultBadge}</td>
            <td class="pe-4 text-muted small" title="${gradesSummary.replace(/"/g, '&quot;')}">${displayGrades}</td>
        `;

        resultsBody.prepend(row);
    }

    // ==========================================
    // 4. COMPREHENSIVE RESULT ANALYSIS LOGIC
    // ==========================================

    // Open Analysis modal trigger handlers
    function openResultAnalysis() {
        try {
            const meta = {
                exam: examSelect ? (examSelect.options[examSelect.selectedIndex]?.text || "HSC") : "HSC",
                year: yearSelect ? (yearSelect.value || "2026") : "2026",
                board: boardSelect ? (boardSelect.options[boardSelect.selectedIndex]?.text || "Board") : "Board"
            };

            if (latestAnalysisData) {
                displayResultAnalysis(latestAnalysisData, meta);
                return;
            }

            if (currentRecords.length > 0) {
                const analysis = computeClientAnalysis(currentRecords);
                latestAnalysisData = analysis;
                displayResultAnalysis(analysis, meta);
                return;
            }

            // If no records are present, offer user to analyze an existing file
            const confirmUpload = confirm(
                "No extracted results are currently loaded on this page.\n\n" +
                "Would you like to select an existing Board Results CSV or Excel file to analyze right now?"
            );
            if (confirmUpload && analysisFileInput) {
                analysisFileInput.click();
            }
        } catch (err) {
            console.error("Error opening result analysis:", err);
            alert(`Analysis Error: ${err.message}`);
        }
    }

    if (openAnalysisBtn) openAnalysisBtn.addEventListener("click", openResultAnalysis);
    if (headerAnalysisBtn) headerAnalysisBtn.addEventListener("click", openResultAnalysis);
    if (bottomAnalysisBtn) bottomAnalysisBtn.addEventListener("click", openResultAnalysis);

    // Navbar integration
    const navOpenAnalysisBtn = document.getElementById("navOpenAnalysisBtn");
    const navBulkLink = document.getElementById("navBulkLink");
    const navIndLink = document.getElementById("navIndLink");
    if (navOpenAnalysisBtn) navOpenAnalysisBtn.addEventListener("click", openResultAnalysis);
    if (navBulkLink) {
        navBulkLink.addEventListener("click", (e) => {
            e.preventDefault();
            const bulkTabBtn = document.getElementById("bulk-tab");
            if (bulkTabBtn) {
                const tab = new bootstrap.Tab(bulkTabBtn);
                tab.show();
            }
            document.getElementById("modeTabs")?.scrollIntoView({ behavior: "smooth" });
        });
    }
    if (navIndLink) {
        navIndLink.addEventListener("click", (e) => {
            e.preventDefault();
            const indTabBtn = document.getElementById("individual-tab");
            if (indTabBtn) {
                const tab = new bootstrap.Tab(indTabBtn);
                tab.show();
            }
            document.getElementById("modeTabs")?.scrollIntoView({ behavior: "smooth" });
        });
    }

    if (uploadFileForAnalysisBtn) {
        uploadFileForAnalysisBtn.addEventListener("click", () => {
            if (analysisFileInput) analysisFileInput.click();
        });
    }

    // 4th Subject Mode Selector State and Listeners
    let currentFourthSubjectMode = "with_4th";
    const btnModeWith4th = document.getElementById("btnModeWith4th");
    const btnModeWithout4th = document.getElementById("btnModeWithout4th");

    if (btnModeWith4th) {
        btnModeWith4th.addEventListener("click", () => {
            currentFourthSubjectMode = "with_4th";
            updateGpaDisplayMode(latestAnalysisData);
        });
    }
    if (btnModeWithout4th) {
        btnModeWithout4th.addEventListener("click", () => {
            currentFourthSubjectMode = "without_4th";
            updateGpaDisplayMode(latestAnalysisData);
        });
    }

    // Upload an existing result file to analyze directly
    if (analysisFileInput) {
        analysisFileInput.addEventListener("change", async () => {
            if (!analysisFileInput.files || analysisFileInput.files.length === 0) return;
            const file = analysisFileInput.files[0];

            const originalBtnHtml = bottomAnalysisBtn.innerHTML;
            bottomAnalysisBtn.disabled = true;
            bottomAnalysisBtn.innerHTML = `<span class="spinner-border spinner-border-sm me-2"></span>Analyzing file...`;

            const formData = new FormData();
            formData.append("file", file);

            try {
                const response = await fetch("/api/analyze-file", {
                    method: "POST",
                    body: formData
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.detail || "Failed to analyze result file.");
                }

                const resJson = await response.json();
                latestAnalysisData = resJson.analysis;

                // Setup Excel Export via direct download stream
                if (exportAnalysisExcelBtn) {
                    exportAnalysisExcelBtn.onclick = async (e) => {
                        e.preventDefault();
                        const fileInstDetails = getInstitutionHeaderDetails(latestAnalysisData, {
                            exam: "Analyzed File",
                            year: "",
                            board: (latestAnalysisData && latestAnalysisData.board) || ""
                        });
                        const exportForm = new FormData();
                        exportForm.append("file", file);
                        if (fileInstDetails.eiin) {
                            exportForm.append("eiin", fileInstDetails.eiin);
                        }
                        const expResp = await fetch("/api/export-analysis-file", {
                            method: "POST",
                            body: exportForm
                        });
                        const blob = await expResp.blob();
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement("a");
                        a.href = url;
                        const fileEiinPrefix = fileInstDetails.eiin ? `${fileInstDetails.eiin}, ` : "";
                        const fileInst = (fileInstDetails.instituteName || "Institution").replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
                        a.download = `${fileEiinPrefix}${fileInst} Board Result Comprehensive Analysis Report - ${file.name.replace(/\.[^/.]+$/, "")}.xlsx`;
                        document.body.appendChild(a);
                        a.click();
                        a.remove();
                    };
                }

                displayResultAnalysis(latestAnalysisData, {
                    exam: "Analyzed File",
                    year: "",
                    board: file.name
                });

            } catch (err) {
                alert(`Analysis Error: ${err.message}`);
            } finally {
                bottomAnalysisBtn.disabled = false;
                bottomAnalysisBtn.innerHTML = originalBtnHtml;
                analysisFileInput.value = "";
            }
        });
    }

    // Display Analysis Modal with All Aggregations
    function displayResultAnalysis(analysisData, metaInfo) {
        if (!analysisData) return;

        const summary = analysisData.summary || {};

        // 1. Overall KPI Cards
        const kpiAppeared = document.getElementById("kpiAppeared");
        const kpiTotalStudents = document.getElementById("kpiTotalStudents");
        const kpiAbsent = document.getElementById("kpiAbsent");
        const kpiAbsentRate = document.getElementById("kpiAbsentRate");
        const kpiPassed = document.getElementById("kpiPassed");
        const kpiPassRate = document.getElementById("kpiPassRate");
        const kpiFailed = document.getElementById("kpiFailed");
        const kpiFailRate = document.getElementById("kpiFailRate");
        const kpiGpa5 = document.getElementById("kpiGpa5");
        const kpiGpa5Rate = document.getElementById("kpiGpa5Rate");
        const kpi1Fail = document.getElementById("kpi1Fail");
        const kpi2Fail = document.getElementById("kpi2Fail");
        const analysisMetaText = document.getElementById("analysisMetaText");

        if (kpiAppeared) kpiAppeared.textContent = summary.appeared || 0;
        if (kpiTotalStudents) kpiTotalStudents.textContent = `${summary.total_students || 0} Processed`;
        if (kpiAbsent) kpiAbsent.textContent = summary.absent || 0;
        if (kpiAbsentRate) kpiAbsentRate.textContent = `${summary.absent_rate || 0}% Absent`;
        if (kpiPassed) kpiPassed.textContent = summary.passed || 0;
        if (kpiPassRate) kpiPassRate.textContent = `${summary.pass_rate || 0}% Pass`;
        if (kpiFailed) kpiFailed.textContent = summary.failed || 0;
        if (kpiFailRate) kpiFailRate.textContent = `${summary.fail_rate || 0}% Fail`;
        if (kpi1Fail) kpi1Fail.textContent = summary.one_subject_fail_count || 0;
        if (kpi2Fail) kpi2Fail.textContent = summary.two_subject_fail_count || 0;

        // 4th Subject Bar Real-time Stats Badges
        const kpiGpa5ValWith = document.getElementById("kpiGpa5ValWith");
        const kpiGpa5RateWith = document.getElementById("kpiGpa5RateWith");
        const kpiGpa5ValWithout = document.getElementById("kpiGpa5ValWithout");
        const kpiGpa5RateWithout = document.getElementById("kpiGpa5RateWithout");

        const countWith = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
        const rateWith = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
        const countWithout = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
        const rateWithout = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;

        if (kpiGpa5ValWith) kpiGpa5ValWith.textContent = countWith;
        if (kpiGpa5RateWith) kpiGpa5RateWith.textContent = `${rateWith}%`;
        if (kpiGpa5ValWithout) kpiGpa5ValWithout.textContent = countWithout;
        if (kpiGpa5RateWithout) kpiGpa5RateWithout.textContent = `${rateWithout}%`;

        const instDetails = getInstitutionHeaderDetails(analysisData, metaInfo);
        if (analysisMetaText) {
            const bLabel = (analysisData && (analysisData.board || (analysisData.summary && analysisData.summary.board))) || metaInfo.board || "Board";
            analysisMetaText.textContent = `${metaInfo.exam || "Examination"} ${metaInfo.year || ""} | ${bLabel} | EIIN: ${instDetails.eiin || "N/A"} | ${summary.total_students || 0} Records Analyzed | ${instDetails.instituteName}`;
        }

        // 2. Regular vs Irregular KPI Quick Banner
        const reg = analysisData.regular_summary || {};
        const irreg = analysisData.irregular_summary || {};

        const kpiRegularBadge = document.getElementById("kpiRegularBadge");
        const kpiRegularRate = document.getElementById("kpiRegularRate");
        const kpiRegularPassed = document.getElementById("kpiRegularPassed");
        const kpiRegularFailed = document.getElementById("kpiRegularFailed");
        const kpiRegularAbsent = document.getElementById("kpiRegularAbsent");
        const kpiRegularGpa5 = document.getElementById("kpiRegularGpa5");
        const kpiRegular1Fail = document.getElementById("kpiRegular1Fail");

        if (kpiRegularBadge) kpiRegularBadge.textContent = `${reg.total || 0} Students`;
        if (kpiRegularRate) kpiRegularRate.textContent = `${reg.pass_rate || 0}% Pass`;
        if (kpiRegularPassed) kpiRegularPassed.textContent = reg.passed || 0;
        if (kpiRegularFailed) kpiRegularFailed.textContent = reg.failed || 0;
        if (kpiRegularAbsent) kpiRegularAbsent.textContent = reg.absent || 0;
        if (kpiRegularGpa5) kpiRegularGpa5.textContent = reg.gpa5_count || 0;
        if (kpiRegular1Fail) kpiRegular1Fail.textContent = reg.one_fail_count || 0;

        const kpiIrregularBadge = document.getElementById("kpiIrregularBadge");
        const kpiIrregularRate = document.getElementById("kpiIrregularRate");
        const kpiIrregularPassed = document.getElementById("kpiIrregularPassed");
        const kpiIrregularFailed = document.getElementById("kpiIrregularFailed");
        const kpiIrregularAbsent = document.getElementById("kpiIrregularAbsent");
        const kpiIrregularGpa5 = document.getElementById("kpiIrregularGpa5");
        const kpiIrregular1Fail = document.getElementById("kpiIrregular1Fail");

        if (kpiIrregularBadge) kpiIrregularBadge.textContent = `${irreg.total || 0} Students`;
        if (kpiIrregularRate) kpiIrregularRate.textContent = `${irreg.pass_rate || 0}% Pass`;
        if (kpiIrregularPassed) kpiIrregularPassed.textContent = irreg.passed || 0;
        if (kpiIrregularFailed) kpiIrregularFailed.textContent = irreg.failed || 0;
        if (kpiIrregularAbsent) kpiIrregularAbsent.textContent = irreg.absent || 0;
        if (kpiIrregularGpa5) kpiIrregularGpa5.textContent = irreg.gpa5_count || 0;
        if (kpiIrregular1Fail) kpiIrregular1Fail.textContent = irreg.one_fail_count || 0;

        // Tab badge counters
        const badge1FailCount = document.getElementById("badge1FailCount");
        const badge2FailCount = document.getElementById("badge2FailCount");
        const badgeMFailCount = document.getElementById("badgeMFailCount");
        const badgeAbsentCount = document.getElementById("badgeAbsentCount");
        const badgeMeritCount = document.getElementById("badgeMeritCount");
        const badgeFailListCount = document.getElementById("badgeFailListCount");

        if (badge1FailCount) badge1FailCount.textContent = (analysisData.one_subject_fail && analysisData.one_subject_fail.total) || 0;
        if (badge2FailCount) badge2FailCount.textContent = (analysisData.two_subject_fail && analysisData.two_subject_fail.total) || 0;
        if (badgeMFailCount) badgeMFailCount.textContent = (analysisData.multi_subject_fail && analysisData.multi_subject_fail.total) || 0;
        const absentTotal = (analysisData.absent_analysis && analysisData.absent_analysis.total_absent !== undefined) ? analysisData.absent_analysis.total_absent : (summary.absent || 0);
        if (badgeAbsentCount) badgeAbsentCount.textContent = absentTotal;
        if (badgeMeritCount) badgeMeritCount.textContent = (analysisData.merit_list && analysisData.merit_list.total) || 0;
        if (badgeFailListCount) badgeFailListCount.textContent = (analysisData.failed_list_subject_wise && analysisData.failed_list_subject_wise.total_failed_students) || 0;

        // Render Executive Failure Summary Banner
        renderFailureSummaryExecutiveBanner(analysisData);

        // Make KPI tiles clickable to jump directly to detailed failure tabs
        if (kpi1Fail) {
            const tile1 = kpi1Fail.closest(".metric-tile");
            if (tile1) {
                tile1.classList.add("clickable-tile");
                tile1.title = "Click to view 1-Subject Fail breakdown";
                tile1.onclick = () => {
                    const tab1 = document.getElementById("tab-1fail");
                    if (tab1) {
                        bootstrap.Tab.getOrCreateInstance(tab1).show();
                        tab1.scrollIntoView({ behavior: "smooth", block: "nearest" });
                    }
                };
            }
        }
        if (kpi2Fail) {
            const tile2 = kpi2Fail.closest(".metric-tile");
            if (tile2) {
                tile2.classList.add("clickable-tile");
                tile2.title = "Click to view 2-Subject Fail breakdown";
                tile2.onclick = () => {
                    const tab2 = document.getElementById("tab-2fail");
                    if (tab2) {
                        bootstrap.Tab.getOrCreateInstance(tab2).show();
                        tab2.scrollIntoView({ behavior: "smooth", block: "nearest" });
                    }
                };
            }
        }

        // Render Tabs
        renderSubjectGradeTableForType("ALL", "all", analysisData.subject_analysis || [], analysisData.available_groups || [], analysisData.available_boards || []);
        renderSubjectGradeTableForType("REGULAR", "reg", analysisData.subject_analysis || [], analysisData.available_groups || [], analysisData.available_boards || []);
        renderSubjectGradeTableForType("IRREGULAR", "irreg", analysisData.subject_analysis || [], analysisData.available_groups || [], analysisData.available_boards || []);
        renderStudentMeritList(analysisData.merit_list || {}, analysisData.available_groups || []);
        renderFailedListSubjectWise(analysisData.failed_list_subject_wise || {}, analysisData.available_groups || []);
        renderOneSubjectFail(analysisData.one_subject_fail || {});
        renderTwoSubjectFail(analysisData.two_subject_fail || {});
        renderMultiSubjectFail(analysisData.multi_subject_fail || {});
        renderGroupAndGpaSummary(analysisData.groups_summary || {}, analysisData.gpa_distribution || {});
        renderAbsentStudents(analysisData.absent_analysis || {});
        renderRegularVsIrregular(analysisData.regular_summary || {}, analysisData.irregular_summary || {}, analysisData.types_summary || {}, analysisData.group_type_comparison || []);

        // Synchronize 4th Subject display mode
        updateGpaDisplayMode(analysisData);

        // Setup Print Button
        if (printAnalysisBtn) {
            printAnalysisBtn.onclick = () => printAnalysisReport(analysisData, metaInfo);
        }

        // Setup Separate Option PDF Download Buttons
        setupSeparateOptionPdfButtons(analysisData, metaInfo);

        // Show Modal
        if (analysisModalInstance) {
            analysisModalInstance.show();
        } else if (analysisModalEl) {
            const bsModal = new bootstrap.Modal(analysisModalEl);
            bsModal.show();
        }
    }

    // Legacy backward-compatibility alias
    function renderSubjectGradeTable(subjectAnalysis, availableGroups, availableBoards) {
        renderSubjectGradeTableForType("ALL", "all", subjectAnalysis, availableGroups, availableBoards);
    }

    // Reusable function for Subject Wise Grade Tables (All, Regular, Irregular)
    function renderSubjectGradeTableForType(targetType, prefix, subjectAnalysis, availableGroups, availableBoards) {
        const groupFilter = document.getElementById(`${prefix}SubjGroupFilter`);
        const boardFilter = document.getElementById(`${prefix}SubjBoardFilter`);
        const searchInput = document.getElementById(`${prefix}SubjSearch`);
        const tableBody = document.getElementById(`${prefix}SubjectGradeBody`) || document.getElementById("subjectGradeBody");
        const countBadge = document.getElementById(`${prefix}SubjCountBadge`);
        const appearedBadge = document.getElementById(`${prefix}SubjAppearedBadge`);
        const passedBadge = document.getElementById(`${prefix}SubjPassedBadge`);
        const passRateBadge = document.getElementById(`${prefix}SubjPassRateBadge`);

        if (!tableBody) return;

        if (groupFilter) {
            const currentVal = groupFilter.value;
            groupFilter.innerHTML = '<option value="ALL">All Groups Combined</option>';
            (availableGroups || []).forEach(grp => {
                const opt = document.createElement("option");
                opt.value = grp;
                opt.textContent = grp;
                if (grp === currentVal) opt.selected = true;
                groupFilter.appendChild(opt);
            });
        }

        if (boardFilter) {
            const currentBoard = boardFilter.value;
            boardFilter.innerHTML = '<option value="ALL">All Boards</option>';
            (availableBoards || []).forEach(b => {
                const opt = document.createElement("option");
                opt.value = b;
                opt.textContent = b;
                if (b === currentBoard) opt.selected = true;
                boardFilter.appendChild(opt);
            });
        }

        function filterAndRender() {
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            tableBody.innerHTML = "";

            const filtered = (subjectAnalysis || []).filter(item => {
                const matchType = (item.type === targetType);
                if (!matchType) return false;
                const matchGroup = (selectedGroup === "ALL") ? (item.group === "ALL") : (item.group === selectedGroup);
                if (!matchGroup) return false;
                if (query && !item.subject.toLowerCase().includes(query)) return false;
                return true;
            });

            // Update KPI summary badges
            const totalSubjs = filtered.length;
            const totalApp = filtered.reduce((acc, it) => acc + (it.appeared || 0), 0);
            const totalPass = filtered.reduce((acc, it) => acc + (it.passed || 0), 0);
            const avgRate = totalApp > 0 ? (totalPass / totalApp * 100).toFixed(2) : "0.00";

            if (countBadge) countBadge.textContent = totalSubjs;
            if (appearedBadge) appearedBadge.textContent = totalApp;
            if (passedBadge) passedBadge.textContent = totalPass;
            if (passRateBadge) passRateBadge.textContent = `${avgRate}%`;

            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="14" class="text-center py-4 text-muted">No subject records found matching the filter.</td></tr>`;
                return;
            }

            filtered.forEach(item => {
                const tr = document.createElement("tr");
                const g = item.grades || {};
                const passPct = item.pass_rate || 0;
                const progressColor = passPct >= 80 ? "bg-success" : (passPct >= 50 ? "bg-primary" : "bg-danger");

                let typeBadge = `<span class="badge bg-secondary-subtle text-secondary border">${item.type}</span>`;
                if (item.type === "REGULAR") {
                    typeBadge = `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`;
                } else if (item.type === "IRREGULAR") {
                    typeBadge = `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`;
                } else if (item.type === "ALL") {
                    typeBadge = `<span class="badge bg-light text-dark border">ALL</span>`;
                }

                tr.innerHTML = `
                    <td class="text-start ps-3 fw-bold text-dark">${item.subject}</td>
                    <td><span class="badge bg-light text-secondary border">${item.group}</span></td>
                    <td>${typeBadge}</td>
                    <td class="fw-semibold">${item.appeared}</td>
                    <td class="text-success fw-bold">${item.passed}</td>
                    <td class="${item.failed > 0 ? 'text-danger fw-bold' : 'text-muted'}">${item.failed}</td>
                    <td>
                        <div class="fw-bold small mb-1">${passPct}%</div>
                        <div class="progress" style="height: 5px;">
                            <div class="progress-bar ${progressColor}" style="width: ${passPct}%"></div>
                        </div>
                    </td>
                    <td class="badge-gpa-gold">${g["A+"] || 0}</td>
                    <td class="badge-gpa-green">${g["A"] || 0}</td>
                    <td class="badge-gpa-green">${g["A-"] || 0}</td>
                    <td class="badge-gpa-blue">${g["B"] || 0}</td>
                    <td class="badge-gpa-blue">${g["C"] || 0}</td>
                    <td>${g["D"] || 0}</td>
                    <td class="${g['F'] > 0 ? 'badge-gpa-red fw-bold text-danger' : 'text-muted'}">${g["F"] || 0}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (groupFilter) groupFilter.onchange = filterAndRender;
        if (boardFilter) boardFilter.onchange = filterAndRender;
        if (searchInput) searchInput.oninput = filterAndRender;
        filterAndRender();
    }

    // TAB 4: Student Merit List
    function renderStudentMeritList(meritData, availableGroups) {
        const tableBody = document.getElementById("meritTableBody");
        if (!tableBody) return;
        const groupFilter = document.getElementById("meritGroupFilter");
        const typeButtonsContainer = document.getElementById("meritTypeButtons");
        const rankModeContainer = document.getElementById("meritRankModeButtons");
        const searchInput = document.getElementById("meritSearch");
        const noticeEl = document.getElementById("meritFilteredNotice");
        const exportBtn = document.getElementById("exportMeritCsvBtn");

        // Summary cards
        const totalPassedEl = document.getElementById("meritTotalPassed");
        const goldenCountEl = document.getElementById("meritGoldenCount");
        const gpa5CountEl = document.getElementById("meritGpa5Count");
        const topStudentNameEl = document.getElementById("meritTopStudentName");
        const topStudentRollEl = document.getElementById("meritTopStudentRoll");

        const students = (meritData && meritData.students) ? meritData.students : [];

        if (totalPassedEl) totalPassedEl.textContent = students.length;
        const goldenCount = students.filter(s => s.Is_GPA5_Without_4th).length;
        const gpa5Count = students.filter(s => s.Is_GPA5_With_4th).length;
        if (goldenCountEl) goldenCountEl.textContent = goldenCount;
        if (gpa5CountEl) gpa5CountEl.textContent = gpa5Count;

        if (students.length > 0) {
            const top1 = students[0];
            if (topStudentNameEl) topStudentNameEl.textContent = top1.Name;
            if (topStudentRollEl) topStudentRollEl.textContent = `Roll: ${top1.Roll} | GPA: ${top1.GPA_With_4th}`;
        } else {
            if (topStudentNameEl) topStudentNameEl.textContent = "None";
            if (topStudentRollEl) topStudentRollEl.textContent = "-";
        }

        // Group Filter
        if (groupFilter) {
            const currentVal = groupFilter.value;
            groupFilter.innerHTML = '<option value="ALL">All Groups</option>';
            (availableGroups || []).forEach(grp => {
                const opt = document.createElement("option");
                opt.value = grp;
                opt.textContent = grp;
                if (grp === currentVal) opt.selected = true;
                groupFilter.appendChild(opt);
            });
        }

        let activeType = "ALL";
        let activeRankMode = "overall"; // "overall" or "group"

        // Type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        // Rank mode toggle buttons
        if (rankModeContainer) {
            rankModeContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeRankMode = btn.getAttribute("data-mode") || "overall";
                    rankModeContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-primary");
                        b.classList.add("btn-outline-primary");
                    });
                    btn.classList.add("active", "btn-primary");
                    btn.classList.remove("btn-outline-primary");
                    filterAndRender();
                };
            });
        }

        let currentlyFiltered = [];

        function filterAndRender() {
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            currentlyFiltered = students.filter(st => {
                if (selectedGroup !== "ALL" && st.Group !== selectedGroup) return false;
                if (activeType !== "ALL" && st.Type !== activeType) return false;
                if (query) {
                    const str = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const groupText = selectedGroup === "ALL" ? "all groups" : selectedGroup;
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${currentlyFiltered.length} ranked students (${groupText}, ${typeText})`;
            }

            if (!tableBody) return;
            tableBody.innerHTML = "";

            if (currentlyFiltered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="10" class="text-center py-4 text-muted">No students found matching merit criteria.</td></tr>`;
                return;
            }

            currentlyFiltered.forEach(st => {
                const tr = document.createElement("tr");

                const rankVal = (activeRankMode === "group") ? st.Group_Merit_Position : st.Merit_Position;
                let rankBadge = `<span class="merit-rank-badge merit-rank-normal">${rankVal}</span>`;
                if (rankVal === 1) rankBadge = `<span class="merit-rank-badge merit-rank-1" title="1st Place (Gold)">🥇 1</span>`;
                else if (rankVal === 2) rankBadge = `<span class="merit-rank-badge merit-rank-2" title="2nd Place (Silver)">🥈 2</span>`;
                else if (rankVal === 3) rankBadge = `<span class="merit-rank-badge merit-rank-3" title="3rd Place (Bronze)">🥉 3</span>`;

                const typeBadge = st.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (st.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

                const gpaWithBadge = st.Is_GPA5_With_4th 
                    ? `<span class="badge badge-gpa-gold px-2.5 py-1.5 fs-7"><i class="bi bi-star-fill me-1"></i>${st.GPA_With_4th}</span>`
                    : `<span class="badge bg-light text-primary border px-2.5 py-1.5 fs-7">${st.GPA_With_4th}</span>`;

                const gpaWithoutBadge = st.Is_GPA5_Without_4th
                    ? `<span class="badge bg-warning text-dark border border-warning px-2 py-1 fs-8" title="Golden A+ (Pure)"><i class="bi bi-award-fill me-1"></i>${st.GPA_Without_4th} (Golden)</span>`
                    : `<span class="badge bg-light text-muted border px-2 py-1 fs-8">${st.GPA_Without_4th}</span>`;

                const gradePills = `
                    <span class="badge bg-warning-subtle text-dark border border-warning-subtle me-1">${st.A_Plus_Count || 0} A+</span>
                    <span class="badge bg-success-subtle text-success border border-success-subtle me-1">${st.A_Count || 0} A</span>
                    <span class="badge bg-info-subtle text-info border border-info-subtle">${st.A_Minus_Count || 0} A-</span>
                `;

                tr.innerHTML = `
                    <td class="text-center">${rankBadge}</td>
                    <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                    <td class="text-secondary small">${st.Registration || "-"}</td>
                    <td class="fw-semibold text-dark">${st.Name}</td>
                    <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                    <td class="text-center">${typeBadge}</td>
                    <td class="text-center">${gpaWithBadge}</td>
                    <td class="text-center">${gpaWithoutBadge}</td>
                    <td class="text-center">${gradePills}</td>
                    <td class="small text-muted text-truncate" style="max-width: 250px;" title="${st.Passed_Grades_Summary}">${st.Passed_Grades_Summary}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        // CSV Export for Merit List
        if (exportBtn) {
            exportBtn.onclick = () => {
                if (!currentlyFiltered || currentlyFiltered.length === 0) {
                    alert("No student records to export.");
                    return;
                }
                const headers = ["Merit Rank", "Group Rank", "Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "A+ Count", "A Count", "Passed Subjects Detail"];
                const rows = currentlyFiltered.map(s => [
                    s.Merit_Position,
                    s.Group_Merit_Position,
                    s.Roll,
                    s.Registration || "",
                    `"${(s.Name || "").replace(/"/g, '""')}"`,
                    s.Group,
                    s.Type || "REGULAR",
                    s.GPA_With_4th,
                    s.GPA_Without_4th,
                    s.A_Plus_Count || 0,
                    s.A_Count || 0,
                    `"${(s.Passed_Grades_Summary || "").replace(/"/g, '""')}"`
                ]);
                const csvContent = "\uFEFF" + [headers.join(","), ...rows.map(r => r.join(","))].join("\r\n");
                const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
                const url = URL.createObjectURL(blob);
                const link = document.createElement("a");
                link.setAttribute("href", url);
                const instDetails = getInstitutionHeaderDetails(analysisData, metaInfo);
                const meritEiinPrefix = instDetails.eiin ? `${instDetails.eiin}, ` : "";
                const meritInst = (instDetails.instituteName || "Institution").replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
                const examLabel = [metaInfo && metaInfo.exam ? metaInfo.exam : "Result", metaInfo && metaInfo.year ? metaInfo.year : ""].filter(Boolean).join(" ").trim();
                link.setAttribute("download", `${meritEiinPrefix}${meritInst} Student Merit List - ${examLabel}.csv`);
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            };
        }

        if (groupFilter) groupFilter.onchange = filterAndRender;
        if (searchInput) searchInput.oninput = filterAndRender;
        filterAndRender();
    }

    // TAB 5: Failed List Subject Wise (Ascending Order)
    function renderFailedListSubjectWise(failedSubjData, availableGroups) {
        const tableBody = document.getElementById("failSubjTableBody");
        if (!tableBody) return;
        const pillsContainer = document.getElementById("failSubjPillsContainer");
        const groupFilter = document.getElementById("failSubjGroupFilter");
        const sortSelect = document.getElementById("failSubjSortOrder");
        const typeButtonsContainer = document.getElementById("failSubjTypeButtons");
        const searchInput = document.getElementById("failSubjSearch");
        const totalBadge = document.getElementById("failSubjTotalBadge");
        const noticeEl = document.getElementById("failSubjFilteredNotice");

        const subjectsList = (failedSubjData && failedSubjData.subjects) ? [...failedSubjData.subjects] : [];
        const entries = (failedSubjData && failedSubjData.entries) ? failedSubjData.entries : [];

        if (totalBadge) totalBadge.textContent = `${entries.length} Subject Failure Entries`;

        if (groupFilter) {
            const currentVal = groupFilter.value;
            groupFilter.innerHTML = '<option value="ALL">All Groups</option>';
            (availableGroups || []).forEach(grp => {
                const opt = document.createElement("option");
                opt.value = grp;
                opt.textContent = grp;
                if (grp === currentVal) opt.selected = true;
                groupFilter.appendChild(opt);
            });
        }

        let activeSubject = "ALL";
        let activeType = "ALL";

        function updateSortAndPills() {
            if (!pillsContainer) return;
            pillsContainer.innerHTML = "";

            const sortMode = sortSelect ? sortSelect.value : "name_asc";
            const sorted = [...subjectsList];

            if (sortMode === "name_asc") {
                sorted.sort((a, b) => a.subject.localeCompare(b.subject));
            } else if (sortMode === "count_asc") {
                sorted.sort((a, b) => a.count - b.count || a.subject.localeCompare(b.subject));
            } else if (sortMode === "count_desc") {
                sorted.sort((a, b) => b.count - a.count || a.subject.localeCompare(b.subject));
            }

            const allBtn = document.createElement("button");
            allBtn.type = "button";
            allBtn.className = (activeSubject === "ALL") ? "btn btn-sm btn-dark filter-pill active" : "btn btn-sm btn-outline-dark filter-pill";
            allBtn.innerHTML = `All Failed Subjects (${entries.length})`;
            allBtn.onclick = () => {
                activeSubject = "ALL";
                updateSortAndPills();
                filterAndRender();
            };
            pillsContainer.appendChild(allBtn);

            sorted.forEach(s => {
                const btn = document.createElement("button");
                btn.type = "button";
                const isAct = (activeSubject === s.subject);
                btn.className = isAct ? "btn btn-sm btn-danger filter-pill active" : "btn btn-sm btn-outline-danger filter-pill";
                btn.innerHTML = `${s.subject} <span class="badge ${isAct ? 'bg-light text-danger' : 'bg-danger'} ms-1">${s.count}</span>`;
                btn.onclick = () => {
                    activeSubject = s.subject;
                    updateSortAndPills();
                    filterAndRender();
                };
                pillsContainer.appendChild(btn);
            });
        }

        // Type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        function filterAndRender() {
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = entries.filter(item => {
                if (activeSubject !== "ALL" && item.Failed_Subject !== activeSubject) return false;
                if (selectedGroup !== "ALL" && item.Group !== selectedGroup) return false;
                if (activeType !== "ALL" && item.Type !== activeType) return false;
                if (query) {
                    const str = `${item.Failed_Subject} ${item.Roll} ${item.Registration || ""} ${item.Name} ${item.Group} ${item.Type || ""}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const subText = activeSubject === "ALL" ? "all subjects" : activeSubject;
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${filtered.length} failed student records in ascending order (${subText}, ${typeText})`;
            }

            if (!tableBody) return;
            tableBody.innerHTML = "";

            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="9" class="text-center py-4 text-muted">No failed records found matching criteria.</td></tr>`;
                return;
            }

            filtered.forEach(item => {
                const tr = document.createElement("tr");

                const typeBadge = item.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (item.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${item.Type || "OTHER"}</span>`);

                const failCountBadge = item.Failed_Count === 1
                    ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">1 Subject Fail</span>`
                    : (item.Failed_Count === 2
                        ? `<span class="badge bg-danger-subtle text-danger border border-danger-subtle">2 Subject Fail</span>`
                        : `<span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle">${item.Failed_Count} Subjects</span>`);

                const allFailsStr = Array.isArray(item.All_Failed_Subjects) ? item.All_Failed_Subjects.join(", ") : (item.All_Failed_Subjects || "");

                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-danger"><i class="bi bi-x-circle-fill me-1"></i>${item.Failed_Subject}</td>
                    <td class="fw-bold text-dark">${item.Roll}</td>
                    <td class="text-secondary small">${item.Registration || "-"}</td>
                    <td class="fw-semibold text-dark">${item.Name}</td>
                    <td><span class="badge bg-light text-secondary border">${item.Group}</span></td>
                    <td>${typeBadge}</td>
                    <td class="text-center">${failCountBadge}</td>
                    <td class="small text-danger fw-semibold text-truncate" style="max-width: 200px;" title="${allFailsStr}">${allFailsStr}</td>
                    <td class="small text-muted text-truncate" style="max-width: 250px;" title="${item.Passed_Grades_Summary}">${item.Passed_Grades_Summary}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (sortSelect) sortSelect.onchange = () => {
            updateSortAndPills();
            filterAndRender();
        };
        if (groupFilter) groupFilter.onchange = filterAndRender;
        if (searchInput) searchInput.oninput = filterAndRender;

        updateSortAndPills();
        filterAndRender();
    }

    // TAB: Regular vs Irregular Performance
    function renderRegularVsIrregular(regularSummary, irregularSummary, typesSummary, groupTypeComparison) {
        const cardRegCount = document.getElementById("cardRegCount");
        const cardRegAppeared = document.getElementById("cardRegAppeared");
        const cardRegPassed = document.getElementById("cardRegPassed");
        const cardRegPassRate = document.getElementById("cardRegPassRate");
        const cardRegFailed = document.getElementById("cardRegFailed");
        const cardRegGpa5 = document.getElementById("cardRegGpa5");
        const cardReg1Fail = document.getElementById("cardReg1Fail");
        const cardReg2Fail = document.getElementById("cardReg2Fail");

        if (cardRegCount) cardRegCount.textContent = `${regularSummary.total || 0} Students`;
        if (cardRegAppeared) cardRegAppeared.textContent = regularSummary.appeared || 0;
        if (cardRegPassed) cardRegPassed.textContent = regularSummary.passed || 0;
        if (cardRegPassRate) cardRegPassRate.textContent = `${regularSummary.pass_rate || 0}%`;
        if (cardRegFailed) cardRegFailed.textContent = regularSummary.failed || 0;
        if (cardRegGpa5) cardRegGpa5.textContent = regularSummary.gpa5_count || 0;
        if (cardReg1Fail) cardReg1Fail.textContent = regularSummary.one_fail_count || 0;
        if (cardReg2Fail) cardReg2Fail.textContent = regularSummary.two_fail_count || 0;

        const cardIrregCount = document.getElementById("cardIrregCount");
        const cardIrregAppeared = document.getElementById("cardIrregAppeared");
        const cardIrregPassed = document.getElementById("cardIrregPassed");
        const cardIrregPassRate = document.getElementById("cardIrregPassRate");
        const cardIrregFailed = document.getElementById("cardIrregFailed");
        const cardIrregGpa5 = document.getElementById("cardIrregGpa5");
        const cardIrreg1Fail = document.getElementById("cardIrreg1Fail");
        const cardIrreg2Fail = document.getElementById("cardIrreg2Fail");

        if (cardIrregCount) cardIrregCount.textContent = `${irregularSummary.total || 0} Students`;
        if (cardIrregAppeared) cardIrregAppeared.textContent = irregularSummary.appeared || 0;
        if (cardIrregPassed) cardIrregPassed.textContent = irregularSummary.passed || 0;
        if (cardIrregPassRate) cardIrregPassRate.textContent = `${irregularSummary.pass_rate || 0}%`;
        if (cardIrregFailed) cardIrregFailed.textContent = irregularSummary.failed || 0;
        if (cardIrregGpa5) cardIrregGpa5.textContent = irregularSummary.gpa5_count || 0;
        if (cardIrreg1Fail) cardIrreg1Fail.textContent = irregularSummary.one_fail_count || 0;
        if (cardIrreg2Fail) cardIrreg2Fail.textContent = irregularSummary.two_fail_count || 0;

        // Render Overall Type Table
        const typeTableBody = document.getElementById("typeSummaryTableBody");
        if (typeTableBody) {
            typeTableBody.innerHTML = "";
            const typeList = Object.values(typesSummary || {});
            if (typeList.length === 0) {
                typeTableBody.innerHTML = `<tr><td colspan="10" class="text-center py-3 text-muted">No candidate type data recorded.</td></tr>`;
            } else {
                typeList.forEach(t => {
                    const badgeClass = t.type === "REGULAR" ? "bg-primary" : (t.type === "IRREGULAR" ? "bg-warning text-dark" : "bg-secondary");
                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td class="text-start ps-3"><span class="badge ${badgeClass} px-2.5 py-1 fw-bold">${t.type}</span></td>
                        <td class="fw-semibold">${t.total}</td>
                        <td>${t.appeared}</td>
                        <td class="text-success fw-bold">${t.passed}</td>
                        <td class="text-danger fw-bold">${t.failed}</td>
                        <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${t.pass_rate}%</span></td>
                        <td class="badge-gpa-gold">${t.gpa5_count}</td>
                        <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${t.one_fail_count}</span></td>
                        <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${t.two_fail_count}</span></td>
                        <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${t.multi_fail_count}</span></td>
                    `;
                    typeTableBody.appendChild(tr);
                });
            }
        }

        // Render Group x Candidate Type Matrix Table
        const groupTypeTableBody = document.getElementById("groupTypeComparisonTableBody");
        if (groupTypeTableBody) {
            groupTypeTableBody.innerHTML = "";
            const matrix = groupTypeComparison || [];
            if (matrix.length === 0) {
                groupTypeTableBody.innerHTML = `<tr><td colspan="11" class="text-center py-3 text-muted">No group x type data recorded.</td></tr>`;
            } else {
                matrix.forEach(row => {
                    const typeBadge = row.type === "REGULAR" 
                        ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">REGULAR</span>`
                        : (row.type === "IRREGULAR" 
                            ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle px-2 py-1">IRREGULAR</span>`
                            : `<span class="badge bg-light text-secondary border px-2 py-1">${row.type}</span>`);

                    const tr = document.createElement("tr");
                    tr.innerHTML = `
                        <td class="text-start ps-3 fw-bold text-dark">${row.group}</td>
                        <td>${typeBadge}</td>
                        <td class="fw-semibold">${row.total}</td>
                        <td>${row.appeared}</td>
                        <td class="text-success fw-bold">${row.passed}</td>
                        <td class="text-danger fw-bold">${row.failed}</td>
                        <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${row.pass_rate}%</span></td>
                        <td class="badge-gpa-gold">${row.gpa5_count}</td>
                        <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${row.one_fail_count}</span></td>
                        <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${row.two_fail_count}</span></td>
                        <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${row.multi_fail_count}</span></td>
                    `;
                    groupTypeTableBody.appendChild(tr);
                });
            }
        }
    }

    // Executive Failure Summary Snapshot Banner
    function renderFailureSummaryExecutiveBanner(analysisData) {
        const oneFail = analysisData.one_subject_fail || {};
        const twoFail = analysisData.two_subject_fail || {};

        const banner1Total = document.getElementById("banner1FailTotal");
        const banner1Count = document.getElementById("banner1FailCount");
        const banner1SubjCount = document.getElementById("banner1FailSubjectCount");
        const banner1TopList = document.getElementById("banner1FailTopList");
        const banner1Badges = document.getElementById("banner1FailBadges");

        const oneTotal = oneFail.total || 0;
        if (banner1Total) banner1Total.textContent = `${oneTotal} Students`;
        if (banner1Count) banner1Count.textContent = oneTotal;

        const oneSubjs = Object.entries(oneFail.subjects_summary || {});
        if (banner1SubjCount) banner1SubjCount.textContent = `${oneSubjs.length} Subjects`;

        if (banner1TopList) {
            if (oneSubjs.length > 0) {
                const topText = oneSubjs.slice(0, 3).map(([s, c]) => `${s} (${c})`).join(", ");
                banner1TopList.textContent = topText;
            } else {
                banner1TopList.textContent = "None (All Passed / No 1-Sub Failures)";
            }
        }

        if (banner1Badges) {
            banner1Badges.innerHTML = "";
            oneSubjs.slice(0, 6).forEach(([subj, count]) => {
                const b = document.createElement("span");
                b.className = "badge bg-warning-subtle text-dark border border-warning-subtle py-1 px-2 fs-8";
                b.innerHTML = `${subj}: <strong class="text-danger">${count}</strong>`;
                banner1Badges.appendChild(b);
            });
            if (oneSubjs.length > 6) {
                const more = document.createElement("span");
                more.className = "badge bg-light text-muted border py-1 px-1.5 fs-8";
                more.textContent = `+${oneSubjs.length - 6} more`;
                banner1Badges.appendChild(more);
            }
        }

        // 2-Subject Fail banner
        const banner2Total = document.getElementById("banner2FailTotal");
        const banner2Count = document.getElementById("banner2FailCount");
        const banner2PairCount = document.getElementById("banner2FailPairCount");
        const banner2TopList = document.getElementById("banner2FailTopList");
        const banner2Badges = document.getElementById("banner2FailBadges");

        const twoTotal = twoFail.total || 0;
        if (banner2Total) banner2Total.textContent = `${twoTotal} Students`;
        if (banner2Count) banner2Count.textContent = twoTotal;

        const twoPairs = Object.entries(twoFail.pair_summary || {});
        const twoSubjs = Object.entries(twoFail.subjects_summary || {});
        if (banner2PairCount) banner2PairCount.textContent = `${twoPairs.length} Pairs`;

        if (banner2TopList) {
            if (twoPairs.length > 0) {
                const topText = twoPairs.slice(0, 2).map(([p, c]) => `${p} (${c})`).join(" | ");
                banner2TopList.textContent = topText;
            } else if (twoSubjs.length > 0) {
                banner2TopList.textContent = twoSubjs.slice(0, 3).map(([s, c]) => `${s} (${c})`).join(", ");
            } else {
                banner2TopList.textContent = "None";
            }
        }

        if (banner2Badges) {
            banner2Badges.innerHTML = "";
            twoPairs.slice(0, 4).forEach(([pair, count]) => {
                const b = document.createElement("span");
                b.className = "badge bg-danger-subtle text-danger border border-danger-subtle py-1 px-2 fs-8";
                b.innerHTML = `${pair}: <strong class="text-danger">${count}</strong>`;
                banner2Badges.appendChild(b);
            });
            if (twoPairs.length > 4) {
                const more = document.createElement("span");
                more.className = "badge bg-light text-muted border py-1 px-1.5 fs-8";
                more.textContent = `+${twoPairs.length - 4} more`;
                banner2Badges.appendChild(more);
            }
        }

        // Jump buttons
        const btnJump1 = document.getElementById("btnJumpTo1Fail");
        if (btnJump1) {
            btnJump1.onclick = () => {
                const tab1 = document.getElementById("tab-1fail");
                if (tab1) {
                    bootstrap.Tab.getOrCreateInstance(tab1).show();
                    tab1.scrollIntoView({ behavior: "smooth", block: "nearest" });
                }
            };
        }
        const btnJump2 = document.getElementById("btnJumpTo2Fail");
        if (btnJump2) {
            btnJump2.onclick = () => {
                const tab2 = document.getElementById("tab-2fail");
                if (tab2) {
                    bootstrap.Tab.getOrCreateInstance(tab2).show();
                    tab2.scrollIntoView({ behavior: "smooth", block: "nearest" });
                }
            };
        }
    }

    // TAB: One Subject Fail with Rich Visual Summary
    function renderOneSubjectFail(oneFailData) {
        const pillsContainer = document.getElementById("oneSubPillsContainer");
        const cardsContainer = document.getElementById("oneSubCardsContainer");
        const tableBody = document.getElementById("oneSubTableBody");
        if (!tableBody) return;

        const summaryBadge = document.getElementById("oneSubTotalSummary");
        const statTotal = document.getElementById("oneSubStatTotal");
        const statTopSubj = document.getElementById("oneSubStatTopSubj");
        const statTopCount = document.getElementById("oneSubStatTopCount");
        const statSubjCount = document.getElementById("oneSubStatSubjCount");
        const statTypeSplit = document.getElementById("oneSubStatTypeSplit");
        const noticeEl = document.getElementById("oneSubFilteredNotice");
        const searchInput = document.getElementById("oneSubSearch");
        const typeButtonsContainer = document.getElementById("oneSubTypeButtons");
        const resetBtn = document.getElementById("oneSubResetFilterBtn");

        const total = oneFailData.total || 0;
        if (summaryBadge) summaryBadge.textContent = `${total} Total Students`;
        if (statTotal) statTotal.textContent = total;

        let activeSubject = "ALL";
        let activeType = "ALL";
        const students = oneFailData.students || [];

        // Candidate Split
        const regCount = students.filter(s => s.Type === "REGULAR").length;
        const irregCount = students.length - regCount;
        if (statTypeSplit) statTypeSplit.textContent = `${regCount} Reg / ${irregCount} Irreg`;

        const subjsEntries = Object.entries(oneFailData.subjects_summary || {});
        if (statSubjCount) statSubjCount.textContent = `${subjsEntries.length} Subjects`;

        if (subjsEntries.length > 0) {
            const [topS, topC] = subjsEntries[0];
            const topPct = total > 0 ? ((topC / total) * 100).toFixed(1) : 0;
            if (statTopSubj) {
                statTopSubj.textContent = topS;
                statTopSubj.title = topS;
            }
            if (statTopCount) statTopCount.textContent = `${topC} Students (${topPct}%)`;
        } else {
            if (statTopSubj) statTopSubj.textContent = "None";
            if (statTopCount) statTopCount.textContent = "0 Students";
        }

        // Render Summary Cards: "How Many Students Failed One Subject and Which Subject"
        if (cardsContainer) {
            cardsContainer.innerHTML = "";
            if (subjsEntries.length === 0) {
                cardsContainer.innerHTML = `<div class="col-12"><div class="p-3 text-center text-muted small bg-light rounded-3">No students failed in only 1 subject.</div></div>`;
            } else {
                subjsEntries.forEach(([subj, count]) => {
                    const pct = total > 0 ? ((count / total) * 100).toFixed(1) : 0;
                    const subjReg = students.filter(s => s.Failed_Subject === subj && s.Type === "REGULAR").length;
                    const subjIrreg = count - subjReg;

                    const col = document.createElement("div");
                    col.className = "col-sm-6 col-md-4 col-lg-3 flex-fill";

                    const card = document.createElement("div");
                    card.className = "p-2.5 rounded-3 subject-fail-card h-100 shadow-xs";
                    card.setAttribute("data-subject", subj);
                    card.innerHTML = `
                        <div class="d-flex justify-content-between align-items-start gap-1 mb-1.5">
                            <span class="fw-bold text-dark text-truncate small" title="${subj}">${subj}</span>
                            <span class="badge bg-danger rounded-pill px-2 py-0.5 fw-bold fs-8">${count} Failed</span>
                        </div>
                        <div class="d-flex justify-content-between align-items-center text-muted mb-1 fs-8">
                            <span>Share: <strong class="text-dark">${pct}%</strong></span>
                            <span class="text-secondary">${subjReg} Reg / ${subjIrreg} Irreg</span>
                        </div>
                        <div class="progress" style="height: 5px; background-color: #f1f5f9;">
                            <div class="progress-bar bg-danger" role="progressbar" style="width: ${Math.max(parseFloat(pct), 6)}%;" aria-valuenow="${pct}" aria-valuemin="0" aria-valuemax="100"></div>
                        </div>
                        <div class="text-end mt-1.5">
                            <span class="text-primary small fw-semibold fs-8"><i class="bi bi-funnel me-1"></i>Filter Students →</span>
                        </div>
                    `;

                    card.onclick = () => {
                        activeSubject = (activeSubject === subj) ? "ALL" : subj;
                        updatePillsAndCards();
                        filterAndRender();
                    };

                    col.appendChild(card);
                    cardsContainer.appendChild(col);
                });
            }
        }

        // Render Filter Pills
        if (pillsContainer) {
            pillsContainer.innerHTML = "";
            const allBtn = document.createElement("button");
            allBtn.type = "button";
            allBtn.className = "btn btn-sm btn-dark filter-pill active";
            allBtn.innerHTML = `All Subjects (${total})`;
            allBtn.onclick = () => {
                activeSubject = "ALL";
                updatePillsAndCards();
                filterAndRender();
            };
            pillsContainer.appendChild(allBtn);

            subjsEntries.forEach(([subj, count]) => {
                const btn = document.createElement("button");
                btn.type = "button";
                btn.className = "btn btn-sm btn-outline-danger filter-pill";
                btn.innerHTML = `${subj} <span class="badge bg-danger ms-1">${count}</span>`;
                btn.onclick = () => {
                    activeSubject = subj;
                    updatePillsAndCards();
                    filterAndRender();
                };
                pillsContainer.appendChild(btn);
            });
        }

        if (resetBtn) {
            resetBtn.onclick = () => {
                activeSubject = "ALL";
                updatePillsAndCards();
                filterAndRender();
            };
        }

        function updatePillsAndCards() {
            // Update cards active border
            if (cardsContainer) {
                cardsContainer.querySelectorAll(".subject-fail-card").forEach(card => {
                    const s = card.getAttribute("data-subject");
                    if (activeSubject !== "ALL" && s === activeSubject) {
                        card.classList.add("active-filter");
                    } else {
                        card.classList.remove("active-filter");
                    }
                });
            }

            // Update pills
            if (pillsContainer) {
                pillsContainer.querySelectorAll(".filter-pill").forEach(btn => {
                    if (activeSubject === "ALL" && btn.textContent.startsWith("All Subjects")) {
                        btn.className = "btn btn-sm btn-dark filter-pill active";
                    } else if (btn.textContent.startsWith(activeSubject)) {
                        btn.className = "btn btn-sm btn-danger filter-pill active";
                    } else {
                        btn.className = btn.textContent.startsWith("All Subjects")
                            ? "btn btn-sm btn-outline-dark filter-pill"
                            : "btn btn-sm btn-outline-danger filter-pill";
                    }
                });
            }
        }

        // Candidate type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        function filterAndRender() {
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                if (activeSubject !== "ALL" && st.Failed_Subject !== activeSubject) return false;
                if (activeType !== "ALL" && st.Type !== activeType) return false;
                if (query) {
                    const text = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${st.Failed_Subject}`.toLowerCase();
                    if (!text.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const subText = activeSubject === "ALL" ? "all subjects" : activeSubject;
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${filtered.length} students (${subText}, ${typeText})`;
            }

            if (!tableBody) return;
            tableBody.innerHTML = "";
            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="8" class="text-center py-4 text-muted">No students found matching criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const typeBadge = st.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (st.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                    <td class="text-secondary small font-monospace">${st.Registration || "-"}</td>
                    <td class="fw-semibold">${st.Name}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                    <td>${typeBadge}</td>
                    <td><span class="badge bg-danger px-2 py-1 fs-7"><i class="bi bi-x-circle me-1"></i>${st.Failed_Subject} (F)</span></td>
                    <td class="small text-muted" title="${st.Passed_Grades_Summary}">${st.Passed_Grades_Summary}</td>
                    <td class="text-center"><span class="badge bg-success-subtle text-success border border-success-subtle">Eligible for Retake</span></td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (searchInput) searchInput.oninput = filterAndRender;
        updatePillsAndCards();
        filterAndRender();
    }

    // TAB: Two Subject Fail with Rich Visual Summary
    function renderTwoSubjectFail(twoFailData) {
        const pillsContainer = document.getElementById("twoSubPillsContainer");
        const freqContainer = document.getElementById("twoSubSubjectFreqContainer");
        const pairsContainer = document.getElementById("twoSubPairsContainer");
        const tableBody = document.getElementById("twoSubTableBody");
        if (!tableBody) return;

        const summaryBadge = document.getElementById("twoSubTotalSummary");
        const statTotal = document.getElementById("twoSubStatTotal");
        const statTopSubj = document.getElementById("twoSubStatTopSubj");
        const statTopCount = document.getElementById("twoSubStatTopCount");
        const statTopPair = document.getElementById("twoSubStatTopPair");
        const statTopPairCount = document.getElementById("twoSubStatTopPairCount");
        const statTypeSplit = document.getElementById("twoSubStatTypeSplit");
        const noticeEl = document.getElementById("twoSubFilteredNotice");
        const searchInput = document.getElementById("twoSubSearch");
        const typeButtonsContainer = document.getElementById("twoSubTypeButtons");
        const resetBtn = document.getElementById("twoSubResetFilterBtn");

        const total = twoFailData.total || 0;
        if (summaryBadge) summaryBadge.textContent = `${total} Total Students`;
        if (statTotal) statTotal.textContent = total;

        let activeSubject = "ALL";
        let activePair = "ALL";
        let activeType = "ALL";
        const students = twoFailData.students || [];

        // Candidate split
        const regCount = students.filter(s => s.Type === "REGULAR").length;
        const irregCount = students.length - regCount;
        if (statTypeSplit) statTypeSplit.textContent = `${regCount} Reg / ${irregCount} Irreg`;

        const subjsEntries = Object.entries(twoFailData.subjects_summary || {});
        const pairsEntries = Object.entries(twoFailData.pair_summary || {});

        // Top Subject
        if (subjsEntries.length > 0) {
            const [topS, topC] = subjsEntries[0];
            if (statTopSubj) {
                statTopSubj.textContent = topS;
                statTopSubj.title = topS;
            }
            if (statTopCount) statTopCount.textContent = `${topC} Occurrences`;
        } else {
            if (statTopSubj) statTopSubj.textContent = "None";
            if (statTopCount) statTopCount.textContent = "0 Occurrences";
        }

        // Top Pair
        if (pairsEntries.length > 0) {
            const [topP, topPC] = pairsEntries[0];
            const pairPct = total > 0 ? ((topPC / total) * 100).toFixed(1) : 0;
            if (statTopPair) {
                statTopPair.textContent = topP;
                statTopPair.title = topP;
            }
            if (statTopPairCount) statTopPairCount.textContent = `${topPC} Students (${pairPct}%)`;
        } else {
            if (statTopPair) statTopPair.textContent = "None";
            if (statTopPairCount) statTopPairCount.textContent = "0 Students";
        }

        // Render Subject Frequency List
        if (freqContainer) {
            freqContainer.innerHTML = "";
            if (subjsEntries.length === 0) {
                freqContainer.innerHTML = `<div class="text-muted small text-center p-2">No 2-subject fail records.</div>`;
            } else {
                subjsEntries.forEach(([subj, count]) => {
                    const pct = total > 0 ? ((count / total) * 100).toFixed(1) : 0;
                    const item = document.createElement("div");
                    item.className = "two-sub-summary-item";
                    item.setAttribute("data-subject", subj);
                    item.innerHTML = `
                        <div class="d-flex justify-content-between align-items-center mb-1">
                            <span class="fw-bold small text-dark text-truncate" title="${subj}">${subj}</span>
                            <span class="badge bg-danger rounded-pill px-2 py-0.5 fs-8">${count} fails (${pct}%)</span>
                        </div>
                        <div class="progress" style="height: 4px; background-color: #f1f5f9;">
                            <div class="progress-bar bg-danger" role="progressbar" style="width: ${Math.max(parseFloat(pct), 6)}%;"></div>
                        </div>
                    `;
                    item.onclick = () => {
                        activePair = "ALL";
                        activeSubject = (activeSubject === subj) ? "ALL" : subj;
                        updateActiveHighlights();
                        filterAndRender();
                    };
                    freqContainer.appendChild(item);
                });
            }
        }

        // Render Pair Combinations List
        if (pairsContainer) {
            pairsContainer.innerHTML = "";
            if (pairsEntries.length === 0) {
                pairsContainer.innerHTML = `<div class="text-muted small text-center p-2">No pair combinations found.</div>`;
            } else {
                pairsEntries.forEach(([pair, count]) => {
                    const pct = total > 0 ? ((count / total) * 100).toFixed(1) : 0;
                    const item = document.createElement("div");
                    item.className = "two-sub-summary-item";
                    item.setAttribute("data-pair", pair);
                    item.innerHTML = `
                        <div class="d-flex justify-content-between align-items-center mb-1">
                            <span class="fw-bold small text-dark text-truncate" title="${pair}"><i class="bi bi-intersect me-1 text-danger"></i>${pair}</span>
                            <span class="badge bg-danger-subtle text-danger border border-danger-subtle rounded-pill px-2 py-0.5 fs-8">${count} (${pct}%)</span>
                        </div>
                        <div class="progress" style="height: 4px; background-color: #f1f5f9;">
                            <div class="progress-bar bg-danger" role="progressbar" style="width: ${Math.max(parseFloat(pct), 6)}%;"></div>
                        </div>
                    `;
                    item.onclick = () => {
                        activeSubject = "ALL";
                        activePair = (activePair === pair) ? "ALL" : pair;
                        updateActiveHighlights();
                        filterAndRender();
                    };
                    pairsContainer.appendChild(item);
                });
            }
        }

        // Render Filter Pills
        if (pillsContainer) {
            pillsContainer.innerHTML = "";
            const allBtn = document.createElement("button");
            allBtn.type = "button";
            allBtn.className = "btn btn-sm btn-dark filter-pill active";
            allBtn.innerHTML = `All Two-Subject Fails (${total})`;
            allBtn.onclick = () => {
                activeSubject = "ALL";
                activePair = "ALL";
                updateActiveHighlights();
                filterAndRender();
            };
            pillsContainer.appendChild(allBtn);

            subjsEntries.forEach(([subj, count]) => {
                const btn = document.createElement("button");
                btn.type = "button";
                btn.className = "btn btn-sm btn-outline-danger filter-pill";
                btn.innerHTML = `${subj} <span class="badge bg-danger ms-1">${count}</span>`;
                btn.onclick = () => {
                    activePair = "ALL";
                    activeSubject = subj;
                    updateActiveHighlights();
                    filterAndRender();
                };
                pillsContainer.appendChild(btn);
            });
        }

        if (resetBtn) {
            resetBtn.onclick = () => {
                activeSubject = "ALL";
                activePair = "ALL";
                updateActiveHighlights();
                filterAndRender();
            };
        }

        function updateActiveHighlights() {
            // Update pills
            if (pillsContainer) {
                pillsContainer.querySelectorAll(".filter-pill").forEach(btn => {
                    if (activeSubject === "ALL" && activePair === "ALL" && btn.textContent.startsWith("All Two-Subject")) {
                        btn.className = "btn btn-sm btn-dark filter-pill active";
                    } else if (activeSubject !== "ALL" && btn.textContent.startsWith(activeSubject)) {
                        btn.className = "btn btn-sm btn-danger filter-pill active";
                    } else {
                        btn.className = btn.textContent.startsWith("All Two-Subject")
                            ? "btn btn-sm btn-outline-dark filter-pill"
                            : "btn btn-sm btn-outline-danger filter-pill";
                    }
                });
            }

            // Update subject freq highlights
            if (freqContainer) {
                freqContainer.querySelectorAll(".two-sub-summary-item").forEach(item => {
                    const s = item.getAttribute("data-subject");
                    if (activeSubject !== "ALL" && s === activeSubject) {
                        item.classList.add("active-filter");
                    } else {
                        item.classList.remove("active-filter");
                    }
                });
            }

            // Update pair highlights
            if (pairsContainer) {
                pairsContainer.querySelectorAll(".two-sub-summary-item").forEach(item => {
                    const p = item.getAttribute("data-pair");
                    if (activePair !== "ALL" && p === activePair) {
                        item.classList.add("active-filter");
                    } else {
                        item.classList.remove("active-filter");
                    }
                });
            }
        }

        // Candidate type toggle buttons
        if (typeButtonsContainer) {
            typeButtonsContainer.querySelectorAll("button").forEach(btn => {
                btn.onclick = () => {
                    activeType = btn.getAttribute("data-type") || "ALL";
                    typeButtonsContainer.querySelectorAll("button").forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        const t = b.getAttribute("data-type");
                        if (t === "ALL") b.className = "btn btn-outline-dark";
                        else if (t === "REGULAR") b.className = "btn btn-outline-primary";
                        else if (t === "IRREGULAR") b.className = "btn btn-outline-warning text-dark";
                    });

                    btn.classList.add("active");
                    if (activeType === "ALL") btn.className = "btn btn-dark active";
                    else if (activeType === "REGULAR") btn.className = "btn btn-primary active";
                    else if (activeType === "IRREGULAR") btn.className = "btn btn-warning text-dark active";

                    filterAndRender();
                };
            });
        }

        function filterAndRender() {
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                const fails = st.Failed_Subjects || [];
                if (activeSubject !== "ALL" && !fails.includes(activeSubject)) return false;
                if (activePair !== "ALL") {
                    const pairParts = activePair.split("&").map(s => s.trim());
                    if (pairParts.length === 2 && (!fails.includes(pairParts[0]) || !fails.includes(pairParts[1]))) {
                        return false;
                    }
                }
                if (activeType !== "ALL" && st.Type !== activeType) return false;
                if (query) {
                    const text = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${fails.join(" ")}`.toLowerCase();
                    if (!text.includes(query)) return false;
                }
                return true;
            });

            if (noticeEl) {
                const subText = activePair !== "ALL" ? activePair : (activeSubject === "ALL" ? "all subjects" : activeSubject);
                const typeText = activeType === "ALL" ? "all types" : activeType;
                noticeEl.textContent = `Showing ${filtered.length} students (${subText}, ${typeText})`;
            }

            if (!tableBody) return;
            tableBody.innerHTML = "";
            if (filtered.length === 0) {
                tableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No students found matching criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const fails = st.Failed_Subjects || [];
                const f1 = fails[0] || "-";
                const f2 = fails[1] || "-";

                const typeBadge = st.Type === "REGULAR" 
                    ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                    : (st.Type === "IRREGULAR" 
                        ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                        : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                    <td class="text-secondary small font-monospace">${st.Registration || "-"}</td>
                    <td class="fw-semibold">${st.Name}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                    <td>${typeBadge}</td>
                    <td>
                        <span class="badge bg-danger px-2 py-1 fs-7 me-1"><i class="bi bi-x-circle me-1"></i>${f1}</span>
                        <span class="badge bg-danger px-2 py-1 fs-7"><i class="bi bi-x-circle me-1"></i>${f2}</span>
                    </td>
                    <td class="small text-muted" title="${st.Passed_Grades_Summary}">${st.Passed_Grades_Summary}</td>
                `;
                tableBody.appendChild(tr);
            });
        }

        if (searchInput) searchInput.oninput = filterAndRender;
        updateActiveHighlights();
        filterAndRender();
    }

    // TAB 4: 3+ Subject Fail
    function renderMultiSubjectFail(multiFailData) {
        const tableBody = document.getElementById("multiSubTableBody");
        const summaryBadge = document.getElementById("multiSubTotalSummary");
        const students = multiFailData.students || [];

        if (summaryBadge) summaryBadge.textContent = `${students.length} Students`;
        if (!tableBody) return;
        tableBody.innerHTML = "";

        if (students.length === 0) {
            tableBody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-muted">No students failed in 3 or more subjects.</td></tr>`;
            return;
        }

        students.forEach(st => {
            const tr = document.createElement("tr");
            const fails = (st.Failed_Subjects || []).join(", ");
            const typeBadge = st.Type === "REGULAR" 
                ? `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`
                : (st.Type === "IRREGULAR" 
                    ? `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`
                    : `<span class="badge bg-light text-secondary border">${st.Type || "OTHER"}</span>`);

            tr.innerHTML = `
                <td class="ps-3 fw-bold text-dark">${st.Roll}</td>
                <td class="fw-semibold">${st.Name}</td>
                <td><span class="badge bg-light text-dark border">${st.Group}</span></td>
                <td>${typeBadge}</td>
                <td class="text-center"><span class="badge bg-danger px-2 py-1">${st.Failed_Count} Subjects</span></td>
                <td class="text-danger small fw-semibold">${fails}</td>
                <td class="small text-muted">${st.Passed_Grades_Summary}</td>
            `;
            tableBody.appendChild(tr);
        });
    }

    // TAB 5: Group Comparison & GPA Breakdown
    function renderGroupAndGpaSummary(groupsSummary, gpaDistribution) {
        const groupTableBody = document.getElementById("groupSummaryTableBody");

        if (groupTableBody) {
            groupTableBody.innerHTML = "";
            Object.values(groupsSummary || {}).forEach(grp => {
                const tr = document.createElement("tr");
                const gpa5With = (grp.gpa5_count_with_4th !== undefined) ? grp.gpa5_count_with_4th : (grp.gpa5_count || 0);
                const gpa5Without = (grp.gpa5_count_without_4th !== undefined) ? grp.gpa5_count_without_4th : 0;
                tr.innerHTML = `
                    <td class="text-start ps-3 fw-bold text-dark">${grp.group}</td>
                    <td class="fw-semibold">${grp.total}</td>
                    <td>${grp.appeared}</td>
                    <td class="text-secondary fw-bold">${grp.absent || 0}</td>
                    <td class="text-success fw-bold">${grp.passed}</td>
                    <td class="text-danger fw-bold">${grp.failed}</td>
                    <td><span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1">${grp.pass_rate}%</span></td>
                    <td class="badge-gpa-gold fw-bold">${gpa5With}</td>
                    <td class="text-primary fw-bold">${gpa5Without}</td>
                    <td><span class="badge bg-primary-subtle text-primary border border-primary-subtle px-2 py-1">${grp.one_fail_count}</span></td>
                    <td><span class="badge bg-danger-subtle text-danger border border-danger-subtle px-2 py-1">${grp.two_fail_count}</span></td>
                    <td><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle px-2 py-1">${grp.multi_fail_count}</span></td>
                `;
                groupTableBody.appendChild(tr);
            });
        }

        renderGpaDistributionTiers(gpaDistribution);
    }

    // Render GPA Distribution Tier Cards
    function renderGpaDistributionTiers(gpaDistribution) {
        const gpaRow = document.getElementById("gpaDistributionRow");
        if (!gpaRow) return;
        gpaRow.innerHTML = "";
        Object.entries(gpaDistribution || {}).forEach(([tier, count]) => {
            const col = document.createElement("div");
            col.className = "col-6 col-md-3 col-lg-auto flex-fill";
            let colorClass = "text-primary";
            if (tier.includes("5.00")) colorClass = "text-warning";
            else if (tier.includes("0.00") || tier.includes("Fail")) colorClass = "text-danger";
            else if (tier.includes("4.00") || tier.includes("3.50")) colorClass = "text-success";

            col.innerHTML = `
                <div class="card border rounded-3 p-3 bg-light text-center h-100">
                    <span class="text-muted fs-8 fw-semibold">${tier}</span>
                    <h4 class="fw-bold ${colorClass} mt-1 mb-0">${count}</h4>
                </div>
            `;
            gpaRow.appendChild(col);
        });
    }

    // Toggle 4th Subject display mode on KPI and GPA cards
    function updateGpaDisplayMode(analysisData) {
        const data = analysisData || latestAnalysisData;
        if (!data) return;

        const summary = data.summary || {};
        const kpiGpa5 = document.getElementById("kpiGpa5");
        const kpiGpa5Rate = document.getElementById("kpiGpa5Rate");
        const kpiGpa5Title = document.getElementById("kpiGpa5Title");
        const gpaDistModeNotice = document.getElementById("gpaDistModeNotice");
        const btnWith = document.getElementById("btnModeWith4th");
        const btnWithout = document.getElementById("btnModeWithout4th");

        if (currentFourthSubjectMode === "with_4th") {
            if (btnWith) {
                btnWith.classList.add("active", "btn-primary");
                btnWith.classList.remove("btn-outline-primary");
            }
            if (btnWithout) {
                btnWithout.classList.remove("active", "btn-primary");
                btnWithout.classList.add("btn-outline-primary");
            }
            if (kpiGpa5Title) kpiGpa5Title.textContent = "GPA 5.0 (With 4th Sub)";
            const countWith = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
            const rateWith = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
            if (kpiGpa5) kpiGpa5.textContent = countWith;
            if (kpiGpa5Rate) kpiGpa5Rate.textContent = `${rateWith}% of Appeared`;
            if (gpaDistModeNotice) gpaDistModeNotice.innerHTML = `<i class="bi bi-info-circle me-1"></i>Showing: With 4th Subject Mode`;
            renderGpaDistributionTiers(data.gpa_distribution_with_4th || data.gpa_distribution || {});
        } else {
            if (btnWith) {
                btnWith.classList.remove("active", "btn-primary");
                btnWith.classList.add("btn-outline-primary");
            }
            if (btnWithout) {
                btnWithout.classList.add("active", "btn-primary");
                btnWithout.classList.remove("btn-outline-primary");
            }
            if (kpiGpa5Title) kpiGpa5Title.textContent = "Golden A+ (Without 4th Sub)";
            const countWithout = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
            const rateWithout = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;
            if (kpiGpa5) kpiGpa5.textContent = countWithout;
            if (kpiGpa5Rate) kpiGpa5Rate.textContent = `${rateWithout}% of Appeared`;
            if (gpaDistModeNotice) gpaDistModeNotice.innerHTML = `<i class="bi bi-star-fill text-warning me-1"></i>Showing: Without 4th Subject (Golden / Pure A+)`;
            renderGpaDistributionTiers(data.gpa_distribution_without_4th || {});
        }
    }

    // TAB 6: Absent Students Analysis & Roster
    function renderAbsentStudents(absentData) {
        if (!absentData) return;

        // 1. KPI Summary Cards in Absent Pane
        const absentPaneTotal = document.getElementById("absentPaneTotal");
        const absentPaneRate = document.getElementById("absentPaneRate");
        const absentPaneRegular = document.getElementById("absentPaneRegular");
        const absentPaneRegularRate = document.getElementById("absentPaneRegularRate");
        const absentPaneIrregular = document.getElementById("absentPaneIrregular");
        const absentPaneIrregularRate = document.getElementById("absentPaneIrregularRate");
        const absentPaneAppeared = document.getElementById("absentPaneAppeared");

        if (absentPaneTotal) absentPaneTotal.textContent = absentData.total_absent || 0;
        if (absentPaneRate) absentPaneRate.textContent = `${absentData.absent_rate || 0}% of Total`;
        if (absentPaneRegular) absentPaneRegular.textContent = absentData.regular_absent || 0;
        if (absentPaneRegularRate) absentPaneRegularRate.textContent = `${absentData.regular_absent_rate || 0}%`;
        if (absentPaneIrregular) absentPaneIrregular.textContent = absentData.irregular_absent || 0;
        if (absentPaneIrregularRate) absentPaneIrregularRate.textContent = `${absentData.irregular_absent_rate || 0}%`;
        if (absentPaneAppeared) absentPaneAppeared.textContent = absentData.appeared || 0;

        // 2. Group-wise Absent Breakdown Table
        const groupTbody = document.getElementById("absentGroupTableBody");
        if (groupTbody) {
            groupTbody.innerHTML = "";
            const groups = Object.values(absentData.groups_absent || {});
            if (groups.length === 0) {
                groupTbody.innerHTML = `<tr><td colspan="5" class="text-center text-muted py-3">No group absent records available.</td></tr>`;
            } else {
                groups.forEach(g => {
                    const tr = document.createElement("tr");
                    const rateColor = g.absent_rate > 5 ? "text-danger fw-bold" : "text-muted";
                    tr.innerHTML = `
                        <td class="text-start ps-3 fw-bold text-dark">${g.group}</td>
                        <td class="fw-semibold">${g.total}</td>
                        <td>${g.appeared}</td>
                        <td class="fw-bold text-danger">${g.absent}</td>
                        <td class="${rateColor}">${g.absent_rate}%</td>
                    `;
                    groupTbody.appendChild(tr);
                });
            }
        }

        // 3. Absent Students Roster Table with filter and search
        const rosterTbody = document.getElementById("absentTableBody");
        const noticeEl = document.getElementById("absentFilteredNotice");
        const searchInput = document.getElementById("absentSearch");
        const typeButtonsContainer = document.getElementById("absentTypeButtons");

        let currentFilterType = "ALL";
        const students = absentData.students || [];

        function filterAndRenderRoster() {
            if (!rosterTbody) return;
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const filtered = students.filter(st => {
                const stType = String(st.Type || "REGULAR").toUpperCase();
                if (currentFilterType !== "ALL" && !stType.includes(currentFilterType)) {
                    return false;
                }
                if (query) {
                    const roll = String(st.Roll || "").toLowerCase();
                    const reg = String(st.Registration || "").toLowerCase();
                    const name = String(st.Name || "").toLowerCase();
                    const father = String(st["Father Name"] || st.Father_Name || "").toLowerCase();
                    const inst = String(st.Institute || "").toLowerCase();
                    const grp = String(st.Group || "").toLowerCase();
                    if (!roll.includes(query) && !reg.includes(query) && !name.includes(query) && !father.includes(query) && !inst.includes(query) && !grp.includes(query)) {
                        return false;
                    }
                }
                return true;
            });

            if (noticeEl) {
                noticeEl.textContent = `Showing ${filtered.length} of ${students.length} absent students`;
            }

            rosterTbody.innerHTML = "";
            if (filtered.length === 0) {
                rosterTbody.innerHTML = `<tr><td colspan="9" class="text-center py-4 text-muted">No absent students match the criteria.</td></tr>`;
                return;
            }

            filtered.forEach(st => {
                const tr = document.createElement("tr");
                let typeBadge = `<span class="badge bg-secondary-subtle text-secondary border">${st.Type || "REGULAR"}</span>`;
                if (st.Type === "REGULAR") {
                    typeBadge = `<span class="badge bg-primary-subtle text-primary border border-primary-subtle">REGULAR</span>`;
                } else if (st.Type === "IRREGULAR") {
                    typeBadge = `<span class="badge bg-warning-subtle text-dark border border-warning-subtle">IRREGULAR</span>`;
                }

                tr.innerHTML = `
                    <td class="ps-3 fw-bold text-dark">${st.Roll || "-"}</td>
                    <td class="font-monospace text-muted small">${st.Registration || "-"}</td>
                    <td class="fw-semibold text-dark">${st.Name || "N/A"}</td>
                    <td class="text-muted small">${st["Father Name"] || st.Father_Name || "-"}</td>
                    <td><span class="badge bg-light text-dark border">${st.Group || "GENERAL"}</span></td>
                    <td>${typeBadge}</td>
                    <td class="small text-truncate" style="max-width: 180px;" title="${st.Institute || ''}">${st.Institute || "-"}</td>
                    <td class="text-center"><span class="badge bg-secondary px-2.5 py-1">ABSENT</span></td>
                    <td class="small text-muted">${st.Remarks || "Did not sit for examination"}</td>
                `;
                rosterTbody.appendChild(tr);
            });
        }

        if (typeButtonsContainer) {
            const btns = typeButtonsContainer.querySelectorAll("button");
            btns.forEach(btn => {
                btn.onclick = () => {
                    btns.forEach(b => {
                        b.classList.remove("active", "btn-dark", "btn-primary", "btn-warning");
                        if (b.dataset.type === "REGULAR") b.classList.add("btn-outline-primary");
                        else if (b.dataset.type === "IRREGULAR") b.classList.add("btn-outline-warning");
                        else b.classList.add("btn-outline-dark");
                    });
                    btn.classList.add("active");
                    if (btn.dataset.type === "REGULAR") {
                        btn.classList.remove("btn-outline-primary");
                        btn.classList.add("btn-primary");
                    } else if (btn.dataset.type === "IRREGULAR") {
                        btn.classList.remove("btn-outline-warning");
                        btn.classList.add("btn-warning");
                    } else {
                        btn.classList.remove("btn-outline-dark");
                        btn.classList.add("btn-dark");
                    }
                    currentFilterType = btn.dataset.type;
                    filterAndRenderRoster();
                };
            });
        }

        if (searchInput) {
            searchInput.oninput = filterAndRenderRoster;
        }

        filterAndRenderRoster();
    }

    // Client-side fallback analyzer (Used if analyzing directly from active memory)
    function computeClientAnalysis(records) {
        const validGrades = ["A+", "A", "A-", "B", "C", "D", "F"];
        const gradePoints = { "A+": 5.0, "A": 4.0, "A-": 3.5, "B": 3.0, "C": 2.0, "D": 1.0, "F": 0.0 };
        const oneFailStudents = [];
        const twoFailStudents = [];
        const multiFailStudents = [];
        const absentStudents = [];
        const meritCandidates = [];
        const allFailedStudents = [];
        const oneSubFreq = {};
        const twoSubFreq = {};
        const subjectStats = {};
        const groupStudents = {};
        const typeStudents = {};

        const gpaDistWith4th = {
            "5.00 (A+)": 0, "4.00 - 4.99 (A)": 0, "3.50 - 3.99 (A-)": 0,
            "3.00 - 3.49 (B)": 0, "2.00 - 2.99 (C)": 0, "1.00 - 1.99 (D)": 0,
            "0.00 (F / Fail)": 0
        };
        const gpaDistWithout4th = {
            "5.00 (A+)": 0, "4.00 - 4.99 (A)": 0, "3.50 - 3.99 (A-)": 0,
            "3.00 - 3.49 (B)": 0, "2.00 - 2.99 (C)": 0, "1.00 - 1.99 (D)": 0,
            "0.00 (F / Fail)": 0
        };

        let passedCount = 0;
        let failedCount = 0;
        let absentCount = 0;

        function identifyFourthSubjectClient(subjectGrades, group) {
            const keys = Object.keys(subjectGrades);
            if (keys.length <= 1) return null;
            const fourthKeywords = [
                "4TH", "FOURTH", "OPTIONAL", "AGRICULTURE", "HIGHER MATH", "BIOLOGY",
                "STATISTICS", "HOME SCIENCE", "PSYCHOLOGY", "GEOGRAPHY", "ISLAMIC STUDIES",
                "CIVICS", "ECONOMICS", "SOCIOLOGY", "SOCIAL WORK", "FINANCE"
            ];
            for (let i = keys.length - 1; i >= 0; i--) {
                const sUpper = keys[i].toUpperCase();
                for (const kw of fourthKeywords) {
                    if (sUpper.includes(kw)) return keys[i];
                }
            }
            if (keys.length >= 7) return keys[keys.length - 1];
            return null;
        }

        records.forEach(raw => {
            const roll = String(raw.Roll || "").trim();
            const reg = String(raw.Registration || "").trim();
            const name = String(raw.Name || "N/A").trim();
            const father = String(raw["Father Name"] || raw.Father_Name || "-").trim();
            const institute = String(raw.Institute || "-").trim();
            let group = String(raw.Group || "GENERAL").trim().toUpperCase();
            if (!group || group === "NAN" || group === "-") group = "GENERAL";

            let resStatus = String(raw["Result Status"] || raw.Result || "").trim().toUpperCase();
            let gpaNum = parseFloat(raw.GPA || 0) || 0.0;

            const subjectGrades = {};
            let rawSubj = raw.Subject_Grades;
            if (typeof rawSubj === "string" && rawSubj.trim().startsWith("{")) {
                try { rawSubj = JSON.parse(rawSubj); } catch (e) {}
            }
            if (rawSubj && typeof rawSubj === "object") {
                Object.entries(rawSubj).forEach(([k, v]) => {
                    const g = String(v).trim().toUpperCase();
                    if (validGrades.includes(g)) subjectGrades[k.trim()] = g;
                });
            } else if (raw.Subject_Grades_Summary || raw["Subject Grades Summary"]) {
                const sumStr = String(raw.Subject_Grades_Summary || raw["Subject Grades Summary"]);
                sumStr.split(",").forEach(part => {
                    if (part.includes(":")) {
                        const [sName, sGrade] = part.split(":");
                        const g = sGrade.trim().toUpperCase();
                        if (validGrades.includes(g)) subjectGrades[sName.trim()] = g;
                    }
                });
            }

            // Detect candidate type
            let cType = String(raw.Type || raw["Candidate Type"] || raw["Student Type"] || "").trim().toUpperCase();
            if (!cType || cType === "NAN" || cType === "-") {
                cType = (Object.keys(subjectGrades).length > 0 && Object.keys(subjectGrades).length <= 4) ? "IRREGULAR" : "REGULAR";
            } else if (cType.includes("IRREG")) {
                cType = "IRREGULAR";
            } else if (cType.includes("REG")) {
                cType = "REGULAR";
            } else if (cType.includes("IMP")) {
                cType = "IMPROVEMENT";
            } else if (cType.includes("PRIV")) {
                cType = "PRIVATE";
            }

            const isAbsent = resStatus.includes("ABSENT") || resStatus.includes("ABS");

            if (isAbsent) {
                absentCount++;
                const absentObj = {
                    Roll: roll,
                    Registration: reg,
                    Name: name,
                    "Father Name": father,
                    Group: group,
                    Type: cType,
                    Institute: institute,
                    Result: "ABSENT",
                    Remarks: "Did not sit for examination",
                    isAbsent: true,
                    isPassed: false,
                    isFail: false
                };
                absentStudents.push(absentObj);

                if (!groupStudents[group]) groupStudents[group] = [];
                groupStudents[group].push(absentObj);

                if (!typeStudents[cType]) typeStudents[cType] = [];
                typeStudents[cType].push(absentObj);
                return;
            }

            // Fourth subject analysis
            const fourthSubj = identifyFourthSubjectClient(subjectGrades, group);
            const mainSubjects = {};
            let fourthGrade = null;
            Object.entries(subjectGrades).forEach(([s, g]) => {
                if (s === fourthSubj) {
                    fourthGrade = g;
                } else {
                    mainSubjects[s] = g;
                }
            });

            const mainFailed = Object.entries(mainSubjects).filter(([_, g]) => g === "F").map(([s]) => s);
            const passedSubjs = Object.entries(subjectGrades).filter(([_, g]) => g !== "F").map(([s, g]) => `${s} (${g})`);
            const failCnt = mainFailed.length; // Failing 4th sub does not fail candidate

            const isFail = resStatus.includes("FAIL") || failCnt > 0;
            const isPassed = !isFail && (gpaNum > 0 || resStatus.includes("PASSED"));

            // Compute GPA without 4th subject
            let gpaWithout4th = 0.0;
            const mainKeys = Object.keys(mainSubjects);
            if (mainKeys.length > 0) {
                const sumMainGp = mainKeys.reduce((acc, s) => acc + (gradePoints[mainSubjects[s]] || 0.0), 0);
                gpaWithout4th = Math.round((sumMainGp / mainKeys.length) * 100) / 100;
            } else {
                gpaWithout4th = gpaNum;
            }

            // Compute GPA with 4th subject
            let computedW4 = gpaNum;
            if (isPassed && mainKeys.length > 0) {
                let fourthBonus = 0.0;
                if (fourthGrade && gradePoints[fourthGrade] !== undefined) {
                    fourthBonus = Math.max(0.0, gradePoints[fourthGrade] - 2.0);
                }
                const sumMainGp = mainKeys.reduce((acc, s) => acc + (gradePoints[mainSubjects[s]] || 0.0), 0);
                computedW4 = Math.min(5.0, Math.round(((sumMainGp + fourthBonus) / mainKeys.length) * 100) / 100);
            }
            const gpaWith4th = (gpaNum > 0) ? Math.max(gpaNum, computedW4) : computedW4;

            const isGpa5With = isPassed && gpaWith4th >= 5.0;
            const isGpa5Without = isPassed && gpaWithout4th >= 5.0;

            if (isPassed) passedCount++;
            else if (isFail) failedCount++;

            const stInfo = {
                Roll: roll,
                Registration: reg,
                Name: name,
                "Father Name": father,
                Institute: institute,
                Group: group,
                Type: cType,
                isPassed,
                isFail,
                isAbsent: false,
                gpaNum: gpaWith4th,
                gpaWith4th: gpaWith4th,
                gpaWithout4th: gpaWithout4th,
                isGpa5With: isGpa5With,
                isGpa5Without: isGpa5Without,
                failCnt,
                Result: isFail ? "FAIL" : (isPassed ? "PASSED" : "UNKNOWN")
            };

            if (!groupStudents[group]) groupStudents[group] = [];
            groupStudents[group].push(stInfo);

            if (!typeStudents[cType]) typeStudents[cType] = [];
            typeStudents[cType].push(stInfo);

            if (isPassed) {
                const aPlusCnt = Object.values(subjectGrades).filter(g => g === "A+").length;
                const aCnt = Object.values(subjectGrades).filter(g => g === "A").length;
                const aMinusCnt = Object.values(subjectGrades).filter(g => g === "A-").length;
                const bCnt = Object.values(subjectGrades).filter(g => g === "B").length;
                const cCnt = Object.values(subjectGrades).filter(g => g === "C").length;
                const dCnt = Object.values(subjectGrades).filter(g => g === "D").length;

                meritCandidates.push({
                    Roll: roll,
                    Registration: reg,
                    Name: name,
                    "Father Name": father,
                    Institute: institute,
                    Group: group,
                    Type: cType,
                    GPA: gpaWith4th.toFixed(2),
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    GPA_With_4th_Numeric: gpaWith4th,
                    GPA_Without_4th_Numeric: gpaWithout4th,
                    Is_GPA5_With_4th: isGpa5With,
                    Is_GPA5_Without_4th: isGpa5Without,
                    A_Plus_Count: aPlusCnt,
                    A_Count: aCnt,
                    A_Minus_Count: aMinusCnt,
                    B_Count: bCnt,
                    C_Count: cCnt,
                    D_Count: dCnt,
                    Subject_Grades: subjectGrades,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
            }

            if (failCnt > 0) {
                allFailedStudents.push({
                    Roll: roll,
                    Registration: reg,
                    Name: name,
                    "Father Name": father,
                    Institute: institute,
                    Group: group,
                    Type: cType,
                    Failed_Count: failCnt,
                    Failed_Subjects: mainFailed,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
            }

            // Fail categories
            if (failCnt === 1) {
                oneFailStudents.push({
                    Roll: roll, Registration: reg, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Subject: mainFailed[0],
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
                oneSubFreq[mainFailed[0]] = (oneSubFreq[mainFailed[0]] || 0) + 1;
            } else if (failCnt === 2) {
                twoFailStudents.push({
                    Roll: roll, Registration: reg, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Subjects: mainFailed,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
                mainFailed.forEach(s => twoSubFreq[s] = (twoSubFreq[s] || 0) + 1);
            } else if (failCnt >= 3) {
                multiFailStudents.push({
                    Roll: roll, Name: name, Group: group, Type: cType,
                    GPA_With_4th: gpaWith4th.toFixed(2),
                    GPA_Without_4th: gpaWithout4th.toFixed(2),
                    Failed_Count: failCnt, Failed_Subjects: mainFailed,
                    Passed_Grades_Summary: passedSubjs.join(", ") || "-"
                });
            }

            // GPA Distribution With 4th
            if (isFail || !isPassed) gpaDistWith4th["0.00 (F / Fail)"]++;
            else if (gpaWith4th >= 5.0) gpaDistWith4th["5.00 (A+)"]++;
            else if (gpaWith4th >= 4.0) gpaDistWith4th["4.00 - 4.99 (A)"]++;
            else if (gpaWith4th >= 3.5) gpaDistWith4th["3.50 - 3.99 (A-)"]++;
            else if (gpaWith4th >= 3.0) gpaDistWith4th["3.00 - 3.49 (B)"]++;
            else if (gpaWith4th >= 2.0) gpaDistWith4th["2.00 - 2.99 (C)"]++;
            else if (gpaWith4th >= 1.0) gpaDistWith4th["1.00 - 1.99 (D)"]++;
            else gpaDistWith4th["0.00 (F / Fail)"]++;

            // GPA Distribution Without 4th
            if (isFail || !isPassed) gpaDistWithout4th["0.00 (F / Fail)"]++;
            else if (gpaWithout4th >= 5.0) gpaDistWithout4th["5.00 (A+)"]++;
            else if (gpaWithout4th >= 4.0) gpaDistWithout4th["4.00 - 4.99 (A)"]++;
            else if (gpaWithout4th >= 3.5) gpaDistWithout4th["3.50 - 3.99 (A-)"]++;
            else if (gpaWithout4th >= 3.0) gpaDistWithout4th["3.00 - 3.49 (B)"]++;
            else if (gpaWithout4th >= 2.0) gpaDistWithout4th["2.00 - 2.99 (C)"]++;
            else if (gpaWithout4th >= 1.0) gpaDistWithout4th["1.00 - 1.99 (D)"]++;
            else gpaDistWithout4th["0.00 (F / Fail)"]++;

            // Subject Grade distribution per (group, type, subject)
            Object.entries(subjectGrades).forEach(([subj, g]) => {
                const keys = [
                    `${group}:::${cType}:::${subj}`,
                    `ALL:::${cType}:::${subj}`,
                    `${group}:::ALL:::${subj}`,
                    `ALL:::ALL:::${subj}`
                ];
                keys.forEach(key => {
                    const [grpPart, typePart, subjPart] = key.split(":::");
                    if (!subjectStats[key]) {
                        subjectStats[key] = { group: grpPart, type: typePart, subject: subjPart, grades: {} };
                        validGrades.forEach(vg => subjectStats[key].grades[vg] = 0);
                    }
                    subjectStats[key].grades[g] = (subjectStats[key].grades[g] || 0) + 1;
                });
            });
        });

        const totalStudents = records.length;
        const appeared = totalStudents - absentCount;
        const passRate = appeared > 0 ? ((passedCount / appeared) * 100).toFixed(2) : "0.00";
        const failRate = appeared > 0 ? ((failedCount / appeared) * 100).toFixed(2) : "0.00";
        const absentRate = totalStudents > 0 ? ((absentCount / totalStudents) * 100).toFixed(2) : "0.00";
        const gpa5CountWith = gpaDistWith4th["5.00 (A+)"];
        const gpa5RateWith = appeared > 0 ? ((gpa5CountWith / appeared) * 100).toFixed(2) : "0.00";
        const gpa5CountWithout = gpaDistWithout4th["5.00 (A+)"];
        const gpa5RateWithout = appeared > 0 ? ((gpa5CountWithout / appeared) * 100).toFixed(2) : "0.00";

        // Build groups summary
        const groupsSummary = {};
        const groupsAbsent = {};
        Object.entries(groupStudents).forEach(([grp, list]) => {
            const grpTot = list.length;
            const grpPass = list.filter(s => s.isPassed).length;
            const grpFail = list.filter(s => s.isFail).length;
            const grpAbs = list.filter(s => s.isAbsent).length;
            const grpApp = grpTot - grpAbs;
            const grpPassRate = grpApp > 0 ? ((grpPass / grpApp) * 100).toFixed(2) : "0.00";
            const grpAbsRate = grpTot > 0 ? ((grpAbs / grpTot) * 100).toFixed(2) : "0.00";
            const grpGpa5With = list.filter(s => s.isGpa5With).length;
            const grpGpa5Without = list.filter(s => s.isGpa5Without).length;
            groupsSummary[grp] = {
                group: grp, total: grpTot, appeared: grpApp, passed: grpPass,
                failed: grpFail, absent: grpAbs, pass_rate: grpPassRate,
                gpa5_count: grpGpa5With,
                gpa5_count_with_4th: grpGpa5With,
                gpa5_count_without_4th: grpGpa5Without,
                one_fail_count: list.filter(s => s.failCnt === 1).length,
                two_fail_count: list.filter(s => s.failCnt === 2).length,
                multi_fail_count: list.filter(s => s.failCnt >= 3).length
            };
            groupsAbsent[grp] = {
                group: grp, total: grpTot, appeared: grpApp, absent: grpAbs, absent_rate: grpAbsRate
            };
        });

        // Build Candidate Types Summary
        const typesSummary = {};
        const typesAbsent = {};
        Object.entries(typeStudents).forEach(([cType, list]) => {
            const tTot = list.length;
            const tPass = list.filter(s => s.isPassed).length;
            const tFail = list.filter(s => s.isFail).length;
            const tAbs = list.filter(s => s.isAbsent).length;
            const tApp = tTot - tAbs;
            const tPassRate = tApp > 0 ? ((tPass / tApp) * 100).toFixed(2) : "0.00";
            const tFailRate = tApp > 0 ? ((tFail / tApp) * 100).toFixed(2) : "0.00";
            const tAbsRate = tTot > 0 ? ((tAbs / tTot) * 100).toFixed(2) : "0.00";
            const tGpa5With = list.filter(s => s.isGpa5With).length;
            const tGpa5Without = list.filter(s => s.isGpa5Without).length;
            typesSummary[cType] = {
                type: cType,
                total: tTot,
                appeared: tApp,
                passed: tPass,
                failed: tFail,
                absent: tAbs,
                pass_rate: tPassRate,
                fail_rate: tFailRate,
                absent_rate: tAbsRate,
                gpa5_count: tGpa5With,
                gpa5_count_with_4th: tGpa5With,
                gpa5_count_without_4th: tGpa5Without,
                gpa5_rate: tApp > 0 ? ((tGpa5With / tApp) * 100).toFixed(2) : "0.00",
                one_fail_count: list.filter(s => s.failCnt === 1).length,
                two_fail_count: list.filter(s => s.failCnt === 2).length,
                multi_fail_count: list.filter(s => s.failCnt >= 3).length
            };
            typesAbsent[cType] = {
                type: cType, total: tTot, appeared: tApp, absent: tAbs, absent_rate: tAbsRate
            };
        });

        // Build Group x Type Matrix
        const groupTypeComparison = [];
        Object.keys(groupStudents).sort().forEach(grp => {
            ["REGULAR", "IRREGULAR", "IMPROVEMENT", "PRIVATE"].forEach(cType => {
                const sub = (groupStudents[grp] || []).filter(s => s.Type === cType);
                if (sub.length > 0) {
                    const subTot = sub.length;
                    const subPass = sub.filter(s => s.isPassed).length;
                    const subFail = sub.filter(s => s.isFail).length;
                    const subAbs = sub.filter(s => s.isAbsent).length;
                    const subApp = subTot - subAbs;
                    const subRate = subApp > 0 ? ((subPass / subApp) * 100).toFixed(2) : "0.00";
                    const subGpa5With = sub.filter(s => s.isGpa5With).length;
                    const subGpa5Without = sub.filter(s => s.isGpa5Without).length;
                    groupTypeComparison.push({
                        group: grp,
                        type: cType,
                        total: subTot,
                        appeared: subApp,
                        passed: subPass,
                        failed: subFail,
                        absent: subAbs,
                        pass_rate: subRate,
                        gpa5_count: subGpa5With,
                        gpa5_count_with_4th: subGpa5With,
                        gpa5_count_without_4th: subGpa5Without,
                        one_fail_count: sub.filter(s => s.failCnt === 1).length,
                        two_fail_count: sub.filter(s => s.failCnt === 2).length,
                        multi_fail_count: sub.filter(s => s.failCnt >= 3).length
                    });
                }
            });
        });

        const regularSummary = typesSummary["REGULAR"] || {
            type: "REGULAR", total: 0, appeared: 0, passed: 0, failed: 0, absent: 0, pass_rate: "0.00",
            gpa5_count: 0, gpa5_count_with_4th: 0, gpa5_count_without_4th: 0,
            one_fail_count: 0, two_fail_count: 0, multi_fail_count: 0
        };
        const irregularSummary = typesSummary["IRREGULAR"] || {
            type: "IRREGULAR", total: 0, appeared: 0, passed: 0, failed: 0, absent: 0, pass_rate: "0.00",
            gpa5_count: 0, gpa5_count_with_4th: 0, gpa5_count_without_4th: 0,
            one_fail_count: 0, two_fail_count: 0, multi_fail_count: 0
        };

        // Build subject analysis list
        const subjectAnalysis = [];
        Object.values(subjectStats).forEach(entry => {
            const tot = Object.values(entry.grades).reduce((a, b) => a + b, 0);
            if (tot === 0) return;
            const fCount = entry.grades["F"] || 0;
            const pCount = tot - fCount;
            const pRate = ((pCount / tot) * 100).toFixed(2);
            subjectAnalysis.push({
                group: entry.group,
                type: entry.type,
                subject: entry.subject,
                appeared: tot,
                passed: pCount,
                failed: fCount,
                pass_rate: pRate,
                grades: entry.grades
            });
        });

        subjectAnalysis.sort((a, b) => a.group.localeCompare(b.group) || a.type.localeCompare(b.type) || a.subject.localeCompare(b.subject));

        const regAbsent = absentStudents.filter(s => s.Type === "REGULAR").length;
        const irregAbsent = absentStudents.filter(s => s.Type === "IRREGULAR").length;

        const absentAnalysis = {
            total_absent: absentCount,
            absent_rate: absentRate,
            regular_absent: regAbsent,
            regular_absent_rate: (regularSummary.total > 0 ? ((regAbsent / regularSummary.total) * 100).toFixed(2) : "0.00"),
            irregular_absent: irregAbsent,
            irregular_absent_rate: (irregularSummary.total > 0 ? ((irregAbsent / irregularSummary.total) * 100).toFixed(2) : "0.00"),
            appeared: appeared,
            groups_absent: groupsAbsent,
            types_absent: typesAbsent,
            students: absentStudents
        };

        // Build Student Merit List
        meritCandidates.sort((a, b) => {
            if (b.GPA_With_4th_Numeric !== a.GPA_With_4th_Numeric) return b.GPA_With_4th_Numeric - a.GPA_With_4th_Numeric;
            if (b.GPA_Without_4th_Numeric !== a.GPA_Without_4th_Numeric) return b.GPA_Without_4th_Numeric - a.GPA_Without_4th_Numeric;
            if (b.A_Plus_Count !== a.A_Plus_Count) return b.A_Plus_Count - a.A_Plus_Count;
            if (b.A_Count !== a.A_Count) return b.A_Count - a.A_Count;
            const rollA = parseInt(a.Roll, 10) || 99999999;
            const rollB = parseInt(b.Roll, 10) || 99999999;
            return rollA - rollB;
        });

        const groupRankCounters = {};
        meritCandidates.forEach((st, idx) => {
            st.Merit_Position = idx + 1;
            groupRankCounters[st.Group] = (groupRankCounters[st.Group] || 0) + 1;
            st.Group_Merit_Position = groupRankCounters[st.Group];
        });

        // Build Failed List Subject Wise (Ascending Order)
        const subjFailMap = {};
        allFailedStudents.forEach(st => {
            (st.Failed_Subjects || []).forEach(subj => {
                if (!subjFailMap[subj]) subjFailMap[subj] = [];
                subjFailMap[subj].push(st);
            });
        });

        const sortedFailSubjs = Object.keys(subjFailMap).sort((a, b) => a.localeCompare(b));
        const failedSubjectsSummary = [];
        const flatFailedEntries = [];

        sortedFailSubjs.forEach(subj => {
            const list = subjFailMap[subj].sort((a, b) => (parseInt(a.Roll, 10) || 0) - (parseInt(b.Roll, 10) || 0));
            failedSubjectsSummary.push({
                subject: subj,
                count: list.length,
                regular_count: list.filter(s => s.Type === "REGULAR").length,
                irregular_count: list.filter(s => s.Type === "IRREGULAR").length,
                students: list
            });
            list.forEach(s => {
                flatFailedEntries.push({
                    Failed_Subject: subj,
                    Roll: s.Roll,
                    Registration: s.Registration,
                    Name: s.Name,
                    Group: s.Group,
                    Type: s.Type,
                    Failed_Count: s.Failed_Count,
                    All_Failed_Subjects: s.Failed_Subjects,
                    Passed_Grades_Summary: s.Passed_Grades_Summary
                });
            });
        });

        return {
            total_students: totalStudents,
            summary: {
                total_students: totalStudents,
                appeared: appeared,
                passed: passedCount,
                failed: failedCount,
                absent: absentCount,
                pass_rate: passRate,
                fail_rate: failRate,
                absent_rate: absentRate,
                gpa_5_count: gpa5CountWith,
                gpa_5_rate: gpa5RateWith,
                gpa_5_count_with_4th: gpa5CountWith,
                gpa_5_rate_with_4th: gpa5RateWith,
                gpa_5_count_without_4th: gpa5CountWithout,
                gpa_5_rate_without_4th: gpa5RateWithout,
                one_subject_fail_count: oneFailStudents.length,
                two_subject_fail_count: twoFailStudents.length,
                multi_subject_fail_count: multiFailStudents.length
            },
            groups_summary: groupsSummary,
            types_summary: typesSummary,
            regular_summary: regularSummary,
            irregular_summary: irregularSummary,
            group_type_comparison: groupTypeComparison,
            subject_analysis: subjectAnalysis,
            merit_list: {
                total: meritCandidates.length,
                students: meritCandidates
            },
            failed_list_subject_wise: {
                total_failed_students: allFailedStudents.length,
                total_subject_failures: flatFailedEntries.length,
                subjects: failedSubjectsSummary,
                entries: flatFailedEntries
            },
            one_subject_fail: {
                total: oneFailStudents.length,
                subjects_summary: Object.fromEntries(Object.entries(oneSubFreq).sort((a, b) => b[1] - a[1])),
                students: oneFailStudents
            },
            two_subject_fail: {
                total: twoFailStudents.length,
                subjects_summary: Object.fromEntries(Object.entries(twoSubFreq).sort((a, b) => b[1] - a[1])),
                students: twoFailStudents
            },
            multi_subject_fail: {
                total: multiFailStudents.length,
                students: multiFailStudents
            },
            absent_analysis: absentAnalysis,
            gpa_distribution: gpaDistWith4th,
            gpa_distribution_with_4th: gpaDistWith4th,
            gpa_distribution_without_4th: gpaDistWithout4th,
            available_groups: Object.keys(groupStudents).sort(),
            available_types: Object.keys(typeStudents).sort(),
            available_boards: []
        };
    }

    // ==========================================
    // INSTITUTION DETAILS RESOLVER (NAME, ADDRESS, EIIN)
    // ==========================================
    function getInstitutionHeaderDetails(analysisData, metaInfo) {
        let rawInstitute = "";
        if (analysisData && analysisData.institute) {
            rawInstitute = analysisData.institute;
        } else if (analysisData && analysisData.summary && analysisData.summary.institute) {
            rawInstitute = analysisData.summary.institute;
        } else if (metaInfo && metaInfo.institute) {
            rawInstitute = metaInfo.institute;
        } else if (analysisData && analysisData.merit_list && analysisData.merit_list.students && analysisData.merit_list.students.length > 0) {
            const found = analysisData.merit_list.students.find(s => s.Institute && s.Institute !== "-" && s.Institute !== "N/A" && s.Institute !== "nan");
            if (found) rawInstitute = found.Institute;
        } else if (analysisData && analysisData.one_subject_fail && analysisData.one_subject_fail.students && analysisData.one_subject_fail.students.length > 0) {
            const found = analysisData.one_subject_fail.students.find(s => s.Institute && s.Institute !== "-" && s.Institute !== "N/A" && s.Institute !== "nan");
            if (found) rawInstitute = found.Institute;
        }

        const dataBoard = (analysisData && (analysisData.board || (analysisData.summary && analysisData.summary.board))) || "";
        const metaBoard = (metaInfo && metaInfo.board && !/\.(csv|xlsx|xls)$/i.test(metaInfo.board)) ? metaInfo.board : "";
        const boardName = dataBoard || metaBoard || "";
        const cleanBoard = boardName.replace(/\s*\(Fast Server\)/i, "").replace(/\s*\(.*\)/, "").trim();
        const bKey = cleanBoard.toLowerCase();

        const KNOWN_INSTITUTIONS_MAP = {
            "132092": { name: "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE", address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "132092" },
            "105820": { name: "COMILLA AJIT GUHA COLLEGE", address: "Chawkbazar, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "105820" },
            "105814": { name: "CUMILLA VICTORIA GOVT. COLLEGE", address: "Victoria College Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "105814" },
            "105786": { name: "CUMILLA ZILLA SCHOOL", address: "Kandirpar, Cumilla Sadar, Cumilla, Bangladesh", eiin: "105786" },
            "107977": { name: "DHAKA COLLEGE", address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh", eiin: "107977" },
            "107975": { name: "DHAKA CITY COLLEGE", address: "Road No 2, Dhanmondi, Dhaka - 1205, Bangladesh", eiin: "107975" },
            "108274": { name: "NOTRE DAME COLLEGE", address: "Toyenbee Circular Road, Motijheel, Dhaka - 1000, Bangladesh", eiin: "108274" },
            "108573": { name: "RAJUK UTTARA MODEL COLLEGE", address: "Sector-6, Uttara Model Town, Dhaka - 1230, Bangladesh", eiin: "108573" },
            "107845": { name: "ADAMJEE CANTONMENT COLLEGE", address: "Dhaka Cantonment, Dhaka - 1206, Bangladesh", eiin: "107845" },
            "108348": { name: "VIQARUNNISA NOON COLLEGE", address: "1/A Bailey Road, Ramna, Dhaka - 1000, Bangladesh", eiin: "108348" },
            "108272": { name: "HOLY CROSS COLLEGE", address: "Tejgaon, Dhaka - 1215, Bangladesh", eiin: "108272" },
            "108035": { name: "GOVT. SCIENCE COLLEGE", address: "Farmgate, Tejgaon, Dhaka - 1215, Bangladesh", eiin: "108035" },
            "132140": { name: "MOHAMMADPUR GOVT. COLLEGE", address: "Satmasjid Road, Mohammadpur, Dhaka - 1207, Bangladesh", eiin: "132140" },
            "133008": { name: "CHITTAGONG COLLEGE", address: "College Road, Chawkbazar, Chattogram - 4203, Bangladesh", eiin: "133008" },
            "126490": { name: "RAJSHAHI COLLEGE", address: "Boalia, Rajshahi - 6000, Bangladesh", eiin: "126490" },
            "100742": { name: "GOVT. BRAJOMOHUN (BM) COLLEGE", address: "College Road, Barishal - 8200, Bangladesh", eiin: "100742" },
            "130456": { name: "MURARI CHAND (MC) COLLEGE", address: "Tilagarh, Sylhet - 3100, Bangladesh", eiin: "130456" },
            "127485": { name: "CARMICHAEL COLLEGE", address: "Lalbag, Rangpur - 5400, Bangladesh", eiin: "127485" },
            "117106": { name: "GOVT. BRAJALAL (BL) COLLEGE", address: "Daulatpur, Khulna - 9202, Bangladesh", eiin: "117106" },
            "111831": { name: "GOVT. ANANDA MOHAN COLLEGE", address: "College Road, Mymensingh - 2200, Bangladesh", eiin: "111831" },
            "106579": { name: "FENI GOVT. COLLEGE", address: "College Road, Feni - 3900, Bangladesh", eiin: "106579" },
            "103284": { name: "BRAHMANBARIA GOVT. COLLEGE", address: "College Road, Brahmanbaria - 3400, Bangladesh", eiin: "103284" },
            "103597": { name: "CHANDPUR GOVT. COLLEGE", address: "Stadium Road, Chandpur - 3600, Bangladesh", eiin: "103597" },
            "107541": { name: "NOAKHALI GOVT. COLLEGE", address: "Maijdee Court, Noakhali - 3800, Bangladesh", eiin: "107541" },
            "107386": { name: "BANDERHAT HIGH SCHOOL & COLLEGE", address: "Banderhat, Lakshmipur, Bangladesh", eiin: "107386" },
            "115957": { name: "GOVT. M. M. COLLEGE, JASHORE", address: "Kharki, Jashore - 7400, Bangladesh", eiin: "115957" },
            "120583": { name: "DINAJPUR GOVT. COLLEGE", address: "Sadar, Dinajpur - 5200, Bangladesh", eiin: "120583" },
            "108156": { name: "GOVT. MADRASA-E-ALIA DHAKA", address: "Bakshibazar, Dhaka - 1211, Bangladesh", eiin: "108156" },
            "108036": { name: "DHAKA POLYTECHNIC INSTITUTE", address: "Tejgaon I/A, Dhaka - 1208, Bangladesh", eiin: "108036" }
        };

        const BOARD_FALLBACKS = {
            "dhaka": { name: "DHAKA COLLEGE", address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh", eiin: "107977" },
            "cumilla": { name: "CUMILLA GOVT. MODEL COLLEGE", address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "105820" },
            "comilla": { name: "CUMILLA GOVT. MODEL COLLEGE", address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "105820" },
            "chittagong": { name: "CHITTAGONG COLLEGE", address: "College Road, Chawkbazar, Chattogram - 4203, Bangladesh", eiin: "133008" },
            "chattogram": { name: "CHITTAGONG COLLEGE", address: "College Road, Chawkbazar, Chattogram - 4203, Bangladesh", eiin: "133008" },
            "rajshahi": { name: "RAJSHAHI COLLEGE", address: "Boalia, Rajshahi - 6000, Bangladesh", eiin: "126490" },
            "barisal": { name: "GOVT. BRAJOMOHUN (BM) COLLEGE", address: "College Road, Barishal - 8200, Bangladesh", eiin: "100742" },
            "barishal": { name: "GOVT. BRAJOMOHUN (BM) COLLEGE", address: "College Road, Barishal - 8200, Bangladesh", eiin: "100742" },
            "sylhet": { name: "MURARI CHAND (MC) COLLEGE", address: "Tilagarh, Sylhet - 3100, Bangladesh", eiin: "130456" },
            "dinajpur": { name: "DINAJPUR GOVT. COLLEGE", address: "Sadar, Dinajpur - 5200, Bangladesh", eiin: "120583" },
            "jessore": { name: "GOVT. M. M. COLLEGE, JASHORE", address: "Kharki, Jashore - 7400, Bangladesh", eiin: "115957" },
            "jashore": { name: "GOVT. M. M. COLLEGE, JASHORE", address: "Kharki, Jashore - 7400, Bangladesh", eiin: "115957" },
            "mymensingh": { name: "GOVT. ANANDA MOHAN COLLEGE", address: "College Road, Mymensingh - 2200, Bangladesh", eiin: "111831" },
            "madrasah": { name: "GOVT. MADRASA-E-ALIA DHAKA", address: "Bakshibazar, Dhaka - 1211, Bangladesh", eiin: "108156" },
            "technical": { name: "DHAKA POLYTECHNIC INSTITUTE", address: "Tejgaon I/A, Dhaka - 1208, Bangladesh", eiin: "108036" }
        };

        const defaultForBoard = BOARD_FALLBACKS[bKey] || null;

        let instituteName = "";
        let eiin = (analysisData && (analysisData.eiin || (analysisData.summary && analysisData.summary.eiin))) || (metaInfo && metaInfo.eiin) || "";
        eiin = (eiin || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(eiin.toLowerCase())) eiin = "";
        let address = (analysisData && (analysisData.address || (analysisData.summary && analysisData.summary.address))) || "";
        address = (address || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(address.toLowerCase())) address = "";

        // 1. Check if rawInstitute has embedded digits like "(105963) ..." or "105963 - ..."
        if (rawInstitute && rawInstitute.trim()) {
            const str = rawInstitute.trim();
            const eiinMatch = str.match(/\(?\b(\d{5,7})\b\)?/);
            if (eiinMatch) {
                if (!eiin) eiin = eiinMatch[1];
                instituteName = str.replace(/\(?\b\d{5,7}\b\)?/g, "").replace(/[-()]/g, " ").replace(/\s+/g, " ").trim();
            } else {
                instituteName = str;
            }
        }

        // 2. Check student records for EIIN if not on summary
        if (!eiin && analysisData) {
            const studentLists = [
                analysisData.merit_list?.students,
                analysisData.one_subject_fail?.students,
                analysisData.two_subject_fail?.students,
                analysisData.absent_analysis?.students
            ];
            for (const list of studentLists) {
                if (Array.isArray(list) && list.length > 0) {
                    const st = list.find(s => s && (s.EIIN || s.eiin) && !["", "nan", "none", "n/a", "-"].includes(String(s.EIIN || s.eiin).trim().toLowerCase()));
                    if (st) {
                        eiin = (st.EIIN || st.eiin).toString().trim();
                        break;
                    }
                }
            }
        }

        // 3. Check per-institution saved EIIN override in localStorage
        if (!eiin && instituteName) {
            try {
                const k = instituteName.toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
                if (k) eiin = localStorage.getItem(`eiin_override::${k}`) || "";
            } catch (e) {}
        }

        // 4. If EIIN is known, obtain canonical address if backend didn't supply one
        if (eiin && KNOWN_INSTITUTIONS_MAP[eiin]) {
            const match = KNOWN_INSTITUTIONS_MAP[eiin];
            return {
                instituteName: instituteName || match.name,
                eiin: eiin,
                address: address || match.address
            };
        }

        // 5. Match by exact known institution name if no EIIN
        const searchTarget = ((instituteName || "") + " " + (rawInstitute || "")).toUpperCase();
        if (!eiin && instituteName) {
            for (const [k, v] of Object.entries(KNOWN_INSTITUTIONS_MAP)) {
                if (searchTarget.includes(v.name)) {
                    eiin = k;
                    if (!address) address = v.address;
                    break;
                }
            }
        }

        // 6. District matching for address
        const bdDistricts = [
            "Bagerhat", "Bandarban", "Barguna", "Barishal", "Barisal", "Bhola", "Bogra", "Bogura", "Brahmanbaria", 
            "Chandpur", "Chattogram", "Chittagong", "Chuadanga", "Comilla", "Cumilla", "Cox's Bazar", 
            "Dhaka", "Dinajpur", "Faridpur", "Feni", "Gaibandha", "Gazipur", "Gopalganj", 
            "Habiganj", "Jamalpur", "Jashore", "Jessore", "Jhalokati", "Jhenaidah", "Joypurhat", 
            "Khagrachhari", "Khulna", "Kishoreganj", "Kurigram", "Kushtia", "Lakshmipur", "Laxmipur", 
            "Lalmonirhat", "Madaripur", "Magura", "Manikganj", "Meherpur", "Moulvibazar", "Munshiganj", 
            "Mymensingh", "Naogaon", "Narail", "Narayanganj", "Narsingdi", "Natore", "Netrokona", 
            "Nilphamari", "Noakhali", "Pabna", "Panchagarh", "Patuakhali", "Pirojpur", "Rajbari", 
            "Rajshahi", "Rangamati", "Rangpur", "Satkhira", "Shariatpur", "Sherpur", "Sirajganj", 
            "Sunamganj", "Sylhet", "Tangail", "Thakurgaon"
        ];
        if (!address && searchTarget) {
            for (const dist of bdDistricts) {
                if (new RegExp(`\\b${dist}\\b`, "i").test(searchTarget)) {
                    const disp = dist.toLowerCase() === "comilla" ? "Cumilla" : (dist.toLowerCase() === "laxmipur" ? "Lakshmipur" : dist);
                    address = `${disp}, Bangladesh`;
                    break;
                }
            }
        }

        // 7. Never assign an unrelated board-default EIIN if an institution name is present
        if (!instituteName && defaultForBoard) {
            instituteName = defaultForBoard.name;
            if (!eiin) eiin = defaultForBoard.eiin;
            if (!address) address = defaultForBoard.address;
        } else if (!address) {
            address = cleanBoard ? `${cleanBoard}, Bangladesh` : "Bangladesh";
        }

        return {
            instituteName: instituteName || (cleanBoard ? `${cleanBoard} Academic Institution` : "Academic Institution"),
            eiin: eiin,
            address: address
        };
    }

    // Print Report in a formatted new window
    function printAnalysisReport(analysisData, metaInfo) {
        const printWin = window.open("", "_blank");
        const instDetails = getInstitutionHeaderDetails(analysisData, metaInfo);
        if (!printWin) {
            alert("Please allow popups to open the print layout.");
            return;
        }

        const summary = analysisData.summary || {};
        const reg = analysisData.regular_summary || {};
        const irreg = analysisData.irregular_summary || {};
        const groupsSummary = analysisData.groups_summary || {};
        const absentData = analysisData.absent_analysis || {};
        const oneFail = analysisData.one_subject_fail || {};
        const twoFail = analysisData.two_subject_fail || {};
        const mFail = analysisData.multi_subject_fail || {};
        const groupType = analysisData.group_type_comparison || [];
        const meritData = analysisData.merit_list || {};
        const failSubjData = analysisData.failed_list_subject_wise || {};

        const gpa5WithCount = (summary.gpa_5_count_with_4th !== undefined) ? summary.gpa_5_count_with_4th : (summary.gpa_5_count || 0);
        const gpa5WithRate = (summary.gpa_5_rate_with_4th !== undefined) ? summary.gpa_5_rate_with_4th : (summary.gpa_5_rate || 0);
        const gpa5WithoutCount = (summary.gpa_5_count_without_4th !== undefined) ? summary.gpa_5_count_without_4th : 0;
        const gpa5WithoutRate = (summary.gpa_5_rate_without_4th !== undefined) ? summary.gpa_5_rate_without_4th : 0;

        // Categorize Subject Analysis into distinct sections
        const subjList = analysisData.subject_analysis || [];
        const allCandidatesSubjs = subjList.filter(it => it.group === "ALL" && it.type === "ALL");
        const regularSubjs = subjList.filter(it => it.group === "ALL" && it.type === "REGULAR");
        const irregularSubjs = subjList.filter(it => it.group === "ALL" && it.type === "IRREGULAR");
        const improvementSubjs = subjList.filter(it => it.group === "ALL" && it.type === "IMPROVEMENT");
        const bizStudiesSubjs = subjList.filter(it => (it.group || "").toUpperCase() === "BUSINESS STUDIES" && it.type === "ALL");
        const humanitiesSubjs = subjList.filter(it => (it.group || "").toUpperCase() === "HUMANITIES" && it.type === "ALL");
        const scienceSubjs = subjList.filter(it => (it.group || "").toUpperCase() === "SCIENCE" && it.type === "ALL");

        const standardGroups = ["ALL", "BUSINESS STUDIES", "HUMANITIES", "SCIENCE"];
        const otherGroupsList = (analysisData.available_groups || []).filter(g => !standardGroups.includes(g.toUpperCase()));

        const failTables = (analysisData.subject_fail_tables && Object.keys(analysisData.subject_fail_tables).length > 0)
            ? analysisData.subject_fail_tables
            : (analysisData.fail_tables || {});
        const availableGroups = analysisData.available_groups || Object.keys(analysisData.groups_summary || {});

        function renderMainPrintFailTable(scopeTitle, categoryTitle, subjDict) {
            const entries = Object.entries(subjDict || {});
            let rows = "";
            let totReg = 0, totIrreg = 0, totAll = 0;
            if (entries.length === 0) {
                rows = `<tr>
                    <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: left; color: #64748b; font-style: italic; padding-left: 8px;">No failures recorded</td>
                    <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #64748b;">0</td>
                    <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #64748b;">0</td>
                    <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #64748b;">0</td>
                </tr>`;
            } else {
                entries.forEach(([s, item]) => {
                    const reg = (item && typeof item === "object") ? (item.regular || 0) : item;
                    const irreg = (item && typeof item === "object") ? (item.irregular || 0) : 0;
                    const tot = (item && typeof item === "object") ? (item.total || (reg + irreg)) : item;
                    totReg += reg; totIrreg += irreg; totAll += tot;

                    rows += `<tr>
                        <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: left; font-weight: 600; color: #0f172a; padding-left: 8px;">${s}</td>
                        <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #1d4ed8; font-weight: 600;">${reg}</td>
                        <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; color: #b45309; font-weight: 600;">${irreg}</td>
                        <td style="border: 1px solid #cbd5e1; padding: 4px 6px; text-align: center; font-weight: 700; color: #b91c1c;">${tot}</td>
                    </tr>`;
                });
            }

            return `
                <div class="fail-main-card" style="page-break-inside: avoid; break-inside: avoid; width: 100%; box-sizing: border-box;">
                    <table class="table-fail-clean" style="width: 100% !important; border: 1.5px solid #cbd5e1 !important; border-collapse: collapse !important; font-size: 8.5pt !important; margin-bottom: 0 !important; background-color: #ffffff !important;">
                        <thead>
                            <tr style="background-color: #f1f5f9;">
                                <th colspan="4" class="fail-main-scope-header" style="border: 1px solid #cbd5e1 !important; padding: 5px 6px !important; text-align: center !important; font-weight: 700 !important; font-size: 9pt !important; color: #0f2942 !important; background-color: #f1f5f9 !important; text-transform: none !important; letter-spacing: normal !important;">${scopeTitle}</th>
                            </tr>
                            <tr style="background-color: #ffffff;">
                                <th colspan="4" class="fail-main-cat-header" style="border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; text-align: center !important; font-weight: 600 !important; font-size: 8pt !important; color: #475569 !important; background-color: #ffffff !important; text-transform: none !important; letter-spacing: normal !important;">${categoryTitle}</th>
                            </tr>
                            <tr style="background-color: #1e3a5f;">
                                <th class="fail-main-col-header" style="border: 1px solid #334155 !important; width: 44%; text-align: left !important; padding: 5px 6px !important; padding-left: 8px !important; font-weight: 700 !important; font-size: 7.8pt !important; color: #ffffff !important; background-color: #1e3a5f !important; text-transform: uppercase !important; letter-spacing: 0.3px !important;">Subject</th>
                                <th class="fail-main-col-header" style="border: 1px solid #334155 !important; width: 18%; text-align: center !important; padding: 5px 6px !important; font-weight: 700 !important; font-size: 7.8pt !important; color: #ffffff !important; background-color: #1e3a5f !important; text-transform: uppercase !important; letter-spacing: 0.3px !important;">Regular</th>
                                <th class="fail-main-col-header" style="border: 1px solid #334155 !important; width: 18%; text-align: center !important; padding: 5px 6px !important; font-weight: 700 !important; font-size: 7.8pt !important; color: #ffffff !important; background-color: #1e3a5f !important; text-transform: uppercase !important; letter-spacing: 0.3px !important;">Irregular</th>
                                <th class="fail-main-col-header" style="border: 1px solid #334155 !important; width: 20%; text-align: center !important; padding: 5px 6px !important; font-weight: 700 !important; font-size: 7.8pt !important; color: #ffffff !important; background-color: #1e3a5f !important; text-transform: uppercase !important; letter-spacing: 0.3px !important;">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${rows}
                            <tr style="background-color: #f8fafc; font-weight: 700; border-top: 1.5px solid #1e3a5f;">
                                <td style="border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; text-align: left !important; padding-left: 8px !important; font-weight: 700; background-color: #f8fafc; color: #0f172a;">Total</td>
                                <td style="border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; text-align: center !important; color: #1d4ed8; font-weight: 700; background-color: #f8fafc;">${totReg}</td>
                                <td style="border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; text-align: center !important; color: #b45309; font-weight: 700; background-color: #f8fafc;">${totIrreg}</td>
                                <td style="border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; text-align: center !important; color: #b91c1c; font-weight: 700; background-color: #f8fafc;">${totAll}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            `;
        }

        function renderReportSubjTable(items, subTitle, badgeText, badgeClass, desc) {
            if (!items || items.length === 0) return "";
            const totApp = items.reduce((a, b) => a + (b.appeared || 0), 0);
            const totPass = items.reduce((a, b) => a + (b.passed || 0), 0);
            const totFail = items.reduce((a, b) => a + (b.failed || 0), 0);
            const passRate = totApp > 0 ? ((totPass / totApp) * 100).toFixed(2) : "0.00";

            return `
                <div class="report-subsection mb-3" style="page-break-inside: avoid; margin-bottom: 20px;">
                    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #1e3a5f; padding-bottom: 5px; margin-bottom: 6px;">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <span class="badge ${badgeClass}" style="font-size: 8pt; padding: 3px 8px; font-weight: 700; text-transform: uppercase;">${badgeText}</span>
                            <strong style="font-size: 10pt; color: #0f2942; text-transform: uppercase; letter-spacing: 0.03em;">${subTitle}</strong>
                        </div>
                        <div style="font-size: 8.5pt; color: #475569;">
                            Subjects: <strong>${items.length}</strong> | Appeared: <strong>${totApp}</strong> | Passed: <strong style="color: #16a34a;">${totPass}</strong> | Failed: <strong style="color: #dc2626;">${totFail}</strong> | Pass Rate: <strong style="color: #2563eb;">${passRate}%</strong>
                        </div>
                    </div>
                    ${desc ? `<div style="font-size: 8.5pt; color: #64748b; margin-bottom: 6px;">${desc}</div>` : ""}
                    <table class="table table-bordered table-sm text-center mb-2" style="font-size: 8.5pt;">
                        <thead>
                            <tr style="background-color: #1e3a5f; color: #ffffff;">
                                <th class="text-start ps-2" style="width: 220px;">Subject Code & Name</th>
                                <th style="width: 60px;">Appeared</th>
                                <th style="width: 55px;">Passed</th>
                                <th style="width: 55px;">Failed</th>
                                <th style="width: 65px;">Pass %</th>
                                <th class="table-highlight-gold" style="width: 42px;">A+</th>
                                <th style="width: 42px;">A</th>
                                <th style="width: 42px;">A-</th>
                                <th style="width: 42px;">B</th>
                                <th style="width: 42px;">C</th>
                                <th style="width: 42px;">D</th>
                                <th class="table-highlight-red" style="width: 55px;">F (Fail)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${items.map(it => `
                                <tr>
                                    <td class="text-start ps-2 fw-semibold text-dark">${it.subject}</td>
                                    <td class="fw-semibold">${it.appeared}</td>
                                    <td class="text-success fw-bold">${it.passed}</td>
                                    <td class="${it.failed > 0 ? 'text-danger fw-bold table-highlight-red' : 'text-muted'}">${it.failed}</td>
                                    <td class="fw-bold ${parseFloat(it.pass_rate) >= 80 ? 'text-success' : (parseFloat(it.pass_rate) < 50 ? 'text-danger' : 'text-primary')}">${it.pass_rate}%</td>
                                    <td class="fw-bold ${(it.grades['A+'] || 0) > 0 ? 'table-highlight-gold' : ''}">${it.grades["A+"] || 0}</td>
                                    <td>${it.grades["A"] || 0}</td>
                                    <td>${it.grades["A-"] || 0}</td>
                                    <td>${it.grades["B"] || 0}</td>
                                    <td>${it.grades["C"] || 0}</td>
                                    <td>${it.grades["D"] || 0}</td>
                                    <td class="${(it.grades['F'] || 0) > 0 ? 'text-danger fw-bold table-highlight-red' : 'text-muted'}">${it.grades["F"] || 0}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>
            `;
        }

        const cleanBoard = (metaInfo && metaInfo.board ? metaInfo.board : "Board").replace(/\s*\(Fast Server\)/i, "").trim();
        const examLabel = [metaInfo && metaInfo.exam ? metaInfo.exam : "Analyzed File", metaInfo && metaInfo.year ? metaInfo.year : ""].filter(Boolean).join(" ").trim();
        const cleanInstName = (instDetails.instituteName || "Institution").replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
        const eiinPrefix = instDetails.eiin ? `${instDetails.eiin}, ` : "";
        const saveReportTitle = `${eiinPrefix}${cleanInstName} Board Result Comprehensive Analysis Report - ${examLabel}`;
        const nowStr = new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' });

        const html = `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>${saveReportTitle}</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
            <style>
                @page {
                    size: A4 portrait;
                    margin: 10mm 9mm 12mm 9mm;
                }
                * {
                    box-sizing: border-box;
                }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                    font-size: 9pt;
                    line-height: 1.35;
                    color: #1e293b;
                    background-color: #f1f5f9;
                    padding: 18px 14px;
                    margin: 0;
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                .report-container {
                    max-width: 1080px;
                    margin: 0 auto;
                    background: #ffffff;
                    padding: 22px 26px;
                    border-radius: 12px;
                    box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.08);
                    border: 1px solid #e2e8f0;
                }
                table.report-master-layout {
                    width: 100% !important;
                    border-collapse: collapse !important;
                    border: none !important;
                    margin: 0 !important;
                    padding: 0 !important;
                    background: transparent !important;
                }
                table.report-master-layout > thead {
                    display: table-header-group !important;
                }
                table.report-master-layout > tfoot {
                    display: table-footer-group !important;
                }
                table.report-master-layout > thead > tr > th,
                table.report-master-layout > thead > tr > td,
                table.report-master-layout > tbody > tr > td,
                table.report-master-layout > tfoot > tr > th,
                table.report-master-layout > tfoot > tr > td {
                    border: none !important;
                    padding: 0 !important;
                    margin: 0 !important;
                    background: transparent !important;
                    font-weight: normal;
                    text-align: left;
                }
                .print-page-header {
                    width: 100%;
                    padding-bottom: 7px;
                    margin-bottom: 11px;
                    text-align: center;
                }
                .header-inst-name {
                    font-size: 16.5pt;
                    font-weight: 800;
                    color: #0f2942;
                    margin: 0 0 3px 0;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    line-height: 1.2;
                }
                .header-inst-meta {
                    font-size: 8.5pt;
                    color: #334155;
                    font-weight: 600;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    gap: 12px;
                    flex-wrap: wrap;
                    margin-top: 3px;
                }
                .meta-dot {
                    color: #94a3b8;
                    font-weight: bold;
                }
                .eiin-tag {
                    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                    font-size: 9pt;
                    font-weight: 800;
                    background: #e2e8f0;
                    color: #0f2942;
                    padding: 1px 7px;
                    border-radius: 4px;
                    letter-spacing: 0.5px;
                    border: 1px solid #cbd5e1;
                    display: inline-block;
                }
                .header-line {
                    height: 2.5px;
                    background: linear-gradient(90deg, #0f2942 0%, #1e3a5f 60%, #3b82f6 100%);
                    margin-top: 8px;
                    border-radius: 2px;
                }
                .print-page-footer {
                    width: 100%;
                    padding-top: 6px;
                    margin-top: 10px;
                }
                .footer-line {
                    height: 1.5px;
                    background: #cbd5e1;
                    margin-bottom: 5px;
                }
                .footer-row {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    font-size: 8.2pt;
                    color: #475569;
                }
                .footer-brand {
                    font-weight: 700;
                    color: #0f2942;
                    letter-spacing: 0.2px;
                }
                .footer-meta {
                    color: #64748b;
                    font-size: 8pt;
                }
                .report-header {
                    border-bottom: 1.5px solid #cbd5e1;
                    padding-bottom: 10px;
                    margin-bottom: 14px;
                }
                .section-title {
                    font-size: 11.5pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.4px;
                    color: #0f2942;
                    border-bottom: 1.5px solid #cbd5e1;
                    padding-bottom: 5px;
                    margin-top: 20px;
                    margin-bottom: 10px;
                    display: flex;
                    align-items: center;
                    gap: 6px;
                }
                .section-desc {
                    font-size: 8.5pt;
                    color: #64748b;
                    margin-top: -6px;
                    margin-bottom: 10px;
                }
                .kpi-card {
                    background: #ffffff;
                    border: 1px solid #e2e8f0;
                    border-radius: 8px;
                    padding: 8px 10px;
                    text-align: center;
                    height: 100%;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
                }
                .kpi-card .kpi-num {
                    font-size: 18px;
                    font-weight: 800;
                    line-height: 1.2;
                    margin-top: 3px;
                    margin-bottom: 2px;
                }
                .kpi-card .kpi-label {
                    font-size: 8pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #64748b;
                }
                .kpi-card .kpi-sub {
                    font-size: 7.5pt;
                    font-weight: 600;
                }
                .kpi-slate { border-top: 3px solid #475569; }
                .kpi-slate .kpi-num { color: #1e293b; }
                .kpi-blue { border-top: 3px solid #2563eb; }
                .kpi-blue .kpi-num { color: #1d4ed8; }
                .kpi-green { border-top: 3px solid #16a34a; }
                .kpi-green .kpi-num { color: #15803d; }
                .kpi-red { border-top: 3px solid #dc2626; }
                .kpi-red .kpi-num { color: #b91c1c; }
                .kpi-gold { border-top: 3px solid #d97706; }
                .kpi-gold .kpi-num { color: #b45309; }
                .kpi-purple { border-top: 3px solid #7c3aed; }
                .kpi-purple .kpi-num { color: #6d28d9; }
                .kpi-orange { border-top: 3px solid #ea580c; }
                .kpi-orange .kpi-num { color: #c2410c; }

                table {
                    width: 100%;
                    margin-bottom: 14px;
                    font-size: 8.5pt;
                    border-collapse: collapse;
                }
                th {
                    background-color: #1e3a5f !important;
                    color: #ffffff !important;
                    font-weight: 600;
                    text-transform: uppercase;
                    font-size: 7.8pt;
                    letter-spacing: 0.3px;
                    padding: 5px 6px;
                    border: 1px solid #334155 !important;
                    text-align: center;
                    vertical-align: middle;
                }
                td {
                    padding: 4.5px 6px;
                    border: 1px solid #cbd5e1;
                    text-align: center;
                    vertical-align: middle;
                    color: #1e293b;
                }
                tbody tr:nth-child(even) { background-color: #f8fafc; }
                tbody tr:hover { background-color: #f1f5f9; }
                .table-highlight-gold { background-color: #fef3c7 !important; font-weight: 700; color: #92400e; border: 1px solid #fde68a !important; }
                .table-highlight-red { background-color: #fee2e2 !important; font-weight: 700; color: #b91c1c; border: 1px solid #fecaca !important; }
                .table-highlight-green { background-color: #dcfce7 !important; font-weight: 700; color: #15803d; border: 1px solid #bbf7d0 !important; }
                .table-fail-clean th.fail-main-scope-header { background-color: #f1f5f9 !important; color: #0f2942 !important; border: 1px solid #cbd5e1 !important; }
                .table-fail-clean th.fail-main-cat-header { background-color: #ffffff !important; color: #475569 !important; border: 1px solid #cbd5e1 !important; }
                .table-fail-clean th.fail-main-col-header { background-color: #1e3a5f !important; color: #ffffff !important; border: 1px solid #334155 !important; }
                .fail-main-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; align-items: start; }
                @media print {
                    .fail-main-grid { display: grid !important; grid-template-columns: repeat(2, 1fr) !important; gap: 8px !important; }
                }
                .guide-box {
                    background-color: #f0fdf4;
                    border: 1px solid #bbf7d0;
                    border-left: 3.5px solid #16a34a;
                    border-radius: 8px;
                    padding: 10px 14px;
                    margin-top: 12px;
                    margin-bottom: 16px;
                }
                .guide-item {
                    font-size: 8.5pt;
                    margin-bottom: 3px;
                    color: #1e3a24;
                }
                .badge-pill-custom {
                    font-size: 8pt;
                    padding: 3px 8px;
                    border-radius: 12px;
                    font-weight: 600;
                    display: inline-block;
                }
                .signoff-section {
                    margin-top: 30px;
                    padding-top: 15px;
                    border-top: 1px dashed #cbd5e1;
                    font-size: 8.5pt;
                    color: #475569;
                    page-break-inside: avoid;
                    break-inside: avoid;
                }
                .sign-line {
                    border-top: 1.5px solid #0f2942;
                    margin: 0 15px;
                    padding-top: 4px;
                    font-weight: 700;
                    color: #0f2942;
                }
                /* ----------------------------------------------------
                   COVER PAGE & INDEX PAGE STYLES (A4 Print-Optimized)
                   ---------------------------------------------------- */
                .pdf-cover-page {
                    background: #ffffff;
                    border: 2.5px solid #0f2942;
                    border-radius: 8px;
                    padding: 24px 22px 18px 22px;
                    margin-bottom: 24px;
                    box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05);
                    position: relative;
                    box-sizing: border-box;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                    min-height: 255mm;
                    page-break-after: always;
                    break-after: page;
                }
                .cover-inner-border {
                    border: 1px solid #cbd5e1;
                    border-radius: 4px;
                    padding: 20px 18px 16px 18px;
                    height: 100%;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                    box-sizing: border-box;
                }
                .cover-seal-wrapper {
                    text-align: center;
                    margin-bottom: 10px;
                }
                .cover-seal-icon {
                    width: 58px;
                    height: 58px;
                    margin: 0 auto 6px auto;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background: linear-gradient(135deg, #0f2942 0%, #1e3a5f 100%);
                    color: #f8fafc;
                    border-radius: 50%;
                    border: 2.5px double #d97706;
                    box-shadow: 0 4px 10px rgba(15, 41, 66, 0.15);
                }
                .cover-org-label {
                    font-size: 8pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 1.2px;
                    color: #475569;
                    margin-bottom: 3px;
                }
                .cover-status-badge {
                    display: inline-block;
                    background: #e0f2fe;
                    color: #0369a1;
                    border: 1px solid #bae6fd;
                    font-size: 7.5pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.6px;
                    padding: 2px 10px;
                    border-radius: 12px;
                }
                .cover-inst-header {
                    text-align: center;
                    margin-top: 6px;
                    margin-bottom: 10px;
                }
                .cover-inst-name {
                    font-size: 19pt;
                    font-weight: 800;
                    color: #0f2942;
                    text-transform: uppercase;
                    letter-spacing: 0.6px;
                    line-height: 1.2;
                    margin: 0 0 5px 0;
                }
                .cover-inst-meta {
                    font-size: 8.8pt;
                    font-weight: 600;
                    color: #334155;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    gap: 10px;
                    flex-wrap: wrap;
                }
                .cover-meta-dot {
                    color: #94a3b8;
                    font-weight: bold;
                }
                .cover-eiin-pill {
                    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                    font-size: 9pt;
                    font-weight: 800;
                    background: #0f2942;
                    color: #ffffff;
                    padding: 1px 7px;
                    border-radius: 4px;
                    letter-spacing: 0.5px;
                }
                .cover-divider-ribbon {
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    margin: 12px 0 14px 0;
                }
                .cover-ribbon-line {
                    flex-grow: 1;
                    height: 2px;
                    background: linear-gradient(90deg, transparent, #cbd5e1, #1e3a5f);
                }
                .cover-ribbon-line.right {
                    background: linear-gradient(90deg, #1e3a5f, #cbd5e1, transparent);
                }
                .cover-ribbon-diamond {
                    margin: 0 10px;
                    color: #d97706;
                    font-size: 11pt;
                }
                .cover-title-section {
                    text-align: center;
                    margin-bottom: 14px;
                }
                .cover-title-badge {
                    display: inline-block;
                    background: #f1f5f9;
                    color: #334155;
                    font-size: 7.8pt;
                    font-weight: 800;
                    letter-spacing: 1.2px;
                    text-transform: uppercase;
                    padding: 2.5px 10px;
                    border-radius: 4px;
                    border: 1px solid #cbd5e1;
                    margin-bottom: 6px;
                }
                .cover-main-title {
                    font-size: 16pt;
                    font-weight: 800;
                    color: #0f2942;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    line-height: 1.25;
                    margin: 0 0 5px 0;
                }
                .cover-exam-badge {
                    display: inline-block;
                    font-size: 10.5pt;
                    font-weight: 700;
                    color: #1d4ed8;
                    background: #eff6ff;
                    border: 1.5px solid #bfdbfe;
                    padding: 3px 14px;
                    border-radius: 6px;
                    margin-top: 3px;
                }
                .cover-subtitle-desc {
                    font-size: 8.2pt;
                    color: #64748b;
                    max-width: 650px;
                    margin: 7px auto 0 auto;
                    line-height: 1.35;
                }
                .cover-kpi-grid {
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 10px;
                    margin: 12px 0;
                }
                .cover-kpi-box {
                    background: #ffffff;
                    border: 1.5px solid #e2e8f0;
                    border-top: 3.5px solid #0f2942;
                    border-radius: 6px;
                    padding: 8px 6px;
                    text-align: center;
                }
                .cover-kpi-box.kpi-blue { border-top-color: #2563eb; }
                .cover-kpi-box.kpi-green { border-top-color: #16a34a; }
                .cover-kpi-box.kpi-gold { border-top-color: #d97706; }
                .cover-kpi-box.kpi-red { border-top-color: #dc2626; }
                .cover-kpi-val {
                    font-size: 17pt;
                    font-weight: 800;
                    line-height: 1.1;
                    color: #0f2942;
                }
                .cover-kpi-lbl {
                    font-size: 7.2pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #64748b;
                    margin-top: 3px;
                }
                .cover-kpi-sub {
                    font-size: 7pt;
                    font-weight: 600;
                    color: #475569;
                    margin-top: 2px;
                }
                .cover-meta-card {
                    background: #f8fafc;
                    border: 1px solid #e2e8f0;
                    border-radius: 6px;
                    padding: 9px 12px;
                    margin-top: 8px;
                    margin-bottom: 10px;
                }
                .cover-meta-grid {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 10px;
                    text-align: left;
                    font-size: 7.8pt;
                    color: #334155;
                }
                .cover-meta-item strong {
                    color: #0f2942;
                    font-size: 7.8pt;
                    display: block;
                    margin-bottom: 1px;
                }
                .cover-footer-note {
                    text-align: center;
                    border-top: 1px dashed #cbd5e1;
                    padding-top: 7px;
                    font-size: 7.2pt;
                    color: #64748b;
                    line-height: 1.35;
                }

                /* Index / TOC Page Styles */
                .pdf-index-page {
                    background: #ffffff;
                    border: 1px solid #cbd5e1;
                    border-radius: 6px;
                    padding: 20px 20px 16px 20px;
                    margin-bottom: 24px;
                    box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05);
                    box-sizing: border-box;
                    display: flex;
                    flex-direction: column;
                    justify-content: space-between;
                    min-height: 255mm;
                    page-break-after: always;
                    break-after: page;
                }
                .index-page-header {
                    border-bottom: 2px solid #0f2942;
                    padding-bottom: 6px;
                    margin-bottom: 10px;
                }
                .index-mini-inst-line {
                    font-size: 7.8pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    color: #64748b;
                    letter-spacing: 0.5px;
                    margin-bottom: 2px;
                }
                .index-title-row {
                    display: flex;
                    justify-content: space-between;
                    align-items: flex-end;
                }
                .index-title-main {
                    font-size: 14pt;
                    font-weight: 800;
                    color: #0f2942;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    margin: 0;
                }
                .index-title-sub {
                    font-size: 8pt;
                    color: #64748b;
                    margin-top: 2px;
                }
                .index-badge-tag {
                    background: #0f2942;
                    color: #ffffff;
                    font-size: 7.5pt;
                    font-weight: 700;
                    padding: 2.5px 8px;
                    border-radius: 4px;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                }
                .index-table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-bottom: 10px;
                    font-size: 8pt;
                }
                .index-table th {
                    background: #1e3a5f !important;
                    color: #ffffff !important;
                    font-size: 7.2pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    padding: 4.5px 7px;
                    border: 1px solid #334155 !important;
                    text-align: left;
                }
                .index-table td {
                    padding: 4px 7px;
                    border: 1px solid #e2e8f0;
                    vertical-align: middle;
                    color: #1e293b;
                    background: #ffffff;
                }
                .index-table tr:nth-child(even) td {
                    background-color: #f8fafc;
                }
                .index-sec-num {
                    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                    font-weight: 800;
                    color: #0f2942;
                    font-size: 8pt;
                    white-space: nowrap;
                }
                .index-sec-title {
                    font-weight: 700;
                    color: #0f2942;
                    font-size: 8.2pt;
                    display: block;
                }
                .index-sec-desc {
                    font-size: 7.2pt;
                    color: #64748b;
                    line-height: 1.25;
                    margin-top: 1px;
                }
                .index-tag-badge {
                    font-size: 6.8pt;
                    font-weight: 700;
                    padding: 1.5px 5px;
                    border-radius: 3px;
                    text-transform: uppercase;
                    letter-spacing: 0.3px;
                    display: inline-block;
                    white-space: nowrap;
                }
                .index-link-btn {
                    display: inline-flex;
                    align-items: center;
                    gap: 3px;
                    font-size: 7.2pt;
                    font-weight: 700;
                    color: #1d4ed8;
                    background: #eff6ff;
                    border: 1px solid #bfdbfe;
                    border-radius: 3px;
                    padding: 1.5px 6px;
                    text-decoration: none;
                    white-space: nowrap;
                }
                .index-link-btn:hover {
                    background: #2563eb;
                    color: #ffffff;
                }
                .index-finder-box {
                    background: #f0fdf4;
                    border: 1.5px solid #bbf7d0;
                    border-left: 3.5px solid #16a34a;
                    border-radius: 6px;
                    padding: 8px 10px;
                    margin-top: 2px;
                    margin-bottom: 8px;
                }
                .index-finder-title {
                    font-size: 8pt;
                    font-weight: 800;
                    color: #15803d;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    margin-bottom: 5px;
                    display: flex;
                    align-items: center;
                    gap: 5px;
                }
                .index-finder-grid {
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 5px 10px;
                }
                .index-finder-item {
                    font-size: 7.5pt;
                    color: #1e3a24;
                    display: flex;
                    align-items: center;
                    gap: 5px;
                    line-height: 1.3;
                }
                .index-finder-badge {
                    font-weight: 800;
                    font-size: 7pt;
                    background: #dcfce7;
                    color: #166534;
                    border: 1px solid #86efac;
                    padding: 1px 4px;
                    border-radius: 3px;
                    white-space: nowrap;
                }
                .index-page-footer {
                    border-top: 1.5px solid #cbd5e1;
                    padding-top: 5px;
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    font-size: 7.2pt;
                    color: #64748b;
                }

                @media print {
                    body { background-color: #ffffff !important; padding: 0 !important; margin: 0 !important; }
                    .report-container { box-shadow: none !important; padding: 0 !important; margin: 0 !important; max-width: 100% !important; border: none !important; border-radius: 0 !important; }
                    .no-print { display: none !important; }
                    .report-section { page-break-inside: auto !important; break-inside: auto !important; }
                    tr { page-break-inside: avoid !important; break-inside: avoid !important; }
                    .kpi-card, .kpi-grid, .fail-main-card, .report-subsection, .signoff-section { page-break-inside: avoid !important; break-inside: avoid !important; }
                    
                    .pdf-cover-page {
                        page-break-after: always !important;
                        break-after: page !important;
                        page-break-inside: avoid !important;
                        break-inside: avoid !important;
                        height: 260mm !important;
                        min-height: 260mm !important;
                        max-height: 265mm !important;
                        border: 2.5px solid #0f2942 !important;
                        border-radius: 0 !important;
                        padding: 14mm 12mm 12mm 12mm !important;
                        margin: 0 !important;
                        box-shadow: none !important;
                        background: #ffffff !important;
                    }
                    .cover-inner-border {
                        border: 1px solid #94a3b8 !important;
                        border-radius: 0 !important;
                        padding: 10mm 10mm 8mm 10mm !important;
                        height: 100% !important;
                        box-sizing: border-box !important;
                        display: flex !important;
                        flex-direction: column !important;
                        justify-content: space-between !important;
                    }
                    .pdf-index-page {
                        page-break-after: always !important;
                        break-after: page !important;
                        page-break-inside: avoid !important;
                        break-inside: avoid !important;
                        height: 260mm !important;
                        min-height: 260mm !important;
                        max-height: 265mm !important;
                        border: 1.5px solid #cbd5e1 !important;
                        border-radius: 0 !important;
                        padding: 12mm 12mm 10mm 12mm !important;
                        margin: 0 !important;
                        box-shadow: none !important;
                        background: #ffffff !important;
                        display: flex !important;
                        flex-direction: column !important;
                        justify-content: space-between !important;
                    }
                    .index-link-btn {
                        background: transparent !important;
                        border: 1px solid #cbd5e1 !important;
                        color: #0f2942 !important;
                        padding: 1px 4px !important;
                    }
                    .index-table td, .index-table th {
                        padding: 3.5px 5px !important;
                    }
                }
            </style>
        </head>
        <body>
            <!-- Interactive Top Action Toolbar (Hidden on paper) -->
            <div class="no-print d-flex justify-content-between align-items-center mb-4 p-3 bg-white rounded-3 shadow-sm border border-slate-200">
                <div class="d-flex align-items-center gap-2">
                    <span class="badge px-3 py-2 text-white" style="background-color: #1e3a5f; font-size: 11px;">
                        <i class="bi bi-file-earmark-pdf-fill me-1"></i> A4 PDF Ready
                    </span>
                    <span class="text-secondary small">Includes Formal Cover Page, Navigation Index & Comprehensive Sections.</span>
                </div>
                <div class="d-flex gap-2">
                    <button class="btn btn-sm px-4 fw-bold text-white shadow-sm" style="background-color: #1e3a5f;" onclick="window.print()">
                        <i class="bi bi-printer-fill me-1"></i> Print / Save as PDF
                    </button>
                    <button class="btn btn-outline-secondary btn-sm px-3" onclick="window.close()">
                        <i class="bi bi-x-lg me-1"></i> Close
                    </button>
                </div>
            </div>

            <div class="report-container">
                <!-- ========================================== -->
                <!-- PAGE 1: FORMAL INSTITUTIONAL COVER PAGE    -->
                <!-- ========================================== -->
                <div class="pdf-cover-page">
                    <div class="cover-inner-border">
                        <!-- Top Institutional Seal & Accreditation Banner -->
                        <div>
                            <div class="cover-seal-wrapper">
                                <div class="cover-seal-icon">
                                    <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                                        <path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"></path>
                                        <path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"></path>
                                        <path d="M12 6v6"></path>
                                        <path d="M9 9h6"></path>
                                    </svg>
                                </div>
                                <div class="cover-org-label">Education Board Result & Academic Performance Analytics</div>
                                <div class="cover-status-badge">
                                    <i class="bi bi-shield-check me-1"></i>Official Institutional Evaluation Dossier &bull; Session ${metaInfo.year || new Date().getFullYear()}
                                </div>
                            </div>

                            <!-- Institution Grand Header -->
                            <div class="cover-inst-header">
                                <h1 class="cover-inst-name">${instDetails.instituteName || "Academic Institution"}</h1>
                                <div class="cover-inst-meta">
                                    <span><i class="bi bi-geo-alt-fill text-danger me-1"></i><strong>Location:</strong> ${instDetails.address || "Bangladesh"}</span>
                                    <span class="cover-meta-dot">&bull;</span>
                                    <span><i class="bi bi-card-heading text-primary me-1"></i><strong>EIIN:</strong> <span class="cover-eiin-pill">${instDetails.eiin || "N/A"}</span></span>
                                    <span class="cover-meta-dot">&bull;</span>
                                    <span><i class="bi bi-award-fill text-success me-1"></i><strong>Board:</strong> ${cleanBoard}</span>
                                </div>
                            </div>

                            <!-- Ribbon Divider -->
                            <div class="cover-divider-ribbon">
                                <div class="cover-ribbon-line"></div>
                                <div class="cover-ribbon-diamond"><i class="bi bi-mortarboard-fill"></i></div>
                                <div class="cover-ribbon-line right"></div>
                            </div>

                            <!-- Main Document Title -->
                            <div class="cover-title-section">
                                <div class="cover-title-badge">Comprehensive Academic Performance Report</div>
                                <h2 class="cover-main-title">ANNUAL BOARD RESULT PERFORMANCE & DIAGNOSTIC REPORT</h2>
                                <div class="cover-exam-badge">
                                    <i class="bi bi-check2-circle me-1"></i>${metaInfo.exam || "Examination"} Examination &mdash; ${metaInfo.year || new Date().getFullYear()}
                                </div>
                                <div class="cover-subtitle-desc">
                                    Official institutional evaluation including subject-level diagnostics, student merit positions, stream-wise comparative analyses, and failure remediation rosters.
                                </div>
                            </div>
                        </div>

                        <!-- Middle: Executive Performance Snapshot (4 KPI Cards) -->
                        <div>
                            <div class="cover-kpi-grid">
                                <div class="cover-kpi-box">
                                    <div class="cover-kpi-val">${summary.total_students || 0}</div>
                                    <div class="cover-kpi-lbl">Total Enrolled</div>
                                    <div class="cover-kpi-sub text-muted">${summary.appeared || 0} Attended &bull; ${summary.absent || 0} Absent</div>
                                </div>
                                <div class="cover-kpi-box kpi-green">
                                    <div class="cover-kpi-val text-success">${summary.pass_rate || 0}%</div>
                                    <div class="cover-kpi-lbl">Overall Pass Rate</div>
                                    <div class="cover-kpi-sub text-success">${summary.passed || 0} Passed (${summary.failed || 0} Failed)</div>
                                </div>
                                <div class="cover-kpi-box kpi-gold">
                                    <div class="cover-kpi-val text-warning">${gpa5WithCount}</div>
                                    <div class="cover-kpi-lbl">GPA 5.0 (With 4th)</div>
                                    <div class="cover-kpi-sub text-muted">${gpa5WithoutCount} Golden A+ (Pure)</div>
                                </div>
                                <div class="cover-kpi-box kpi-red">
                                    <div class="cover-kpi-val text-danger">${summary.one_subject_fail_count || 0}</div>
                                    <div class="cover-kpi-lbl">1-Subject Fail</div>
                                    <div class="cover-kpi-sub text-danger">Supplementary Eligible</div>
                                </div>
                            </div>
                        </div>

                        <!-- Bottom: Metadata, Certification & Formal Imprint -->
                        <div>
                            <div class="cover-meta-card">
                                <div class="cover-meta-grid">
                                    <div class="cover-meta-item">
                                        <strong>DOCUMENT STATUS</strong>
                                        <span>Official & Verified Copy</span>
                                    </div>
                                    <div class="cover-meta-item">
                                        <strong>REPORT GENERATION</strong>
                                        <span>${nowStr}</span>
                                    </div>
                                    <div class="cover-meta-item">
                                        <strong>CURRICULUM STREAMS</strong>
                                        <span>Science, Humanities, Business Studies</span>
                                    </div>
                                </div>
                            </div>

                            <div class="cover-footer-note">
                                <div class="fw-semibold text-dark mb-1">
                                    <i class="bi bi-shield-lock-fill me-1 text-primary"></i> CONFIDENTIAL INSTITUTIONAL RECORD &bull; FOR INTERNAL & ACADEMIC AUDIT USE ONLY
                                </div>
                                <div>Generated By iCT Center &nbsp;&bull;&nbsp; Contact: 01641711531 &nbsp;&bull;&nbsp; Page 1 (Cover Page)</div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- ========================================== -->
                <!-- PAGE 2: REPORT INDEX / TABLE OF CONTENTS   -->
                <!-- ========================================== -->
                <div class="pdf-index-page">
                    <div>
                        <!-- Mini Header -->
                        <div class="index-page-header">
                            <div class="index-mini-inst-line">
                                ${instDetails.instituteName || "Institution"} &bull; EIIN: ${instDetails.eiin || "N/A"} &bull; ${cleanBoard} Board &bull; ${metaInfo.exam || "Exam"} ${metaInfo.year || ""}
                            </div>
                            <div class="index-title-row">
                                <div>
                                    <h2 class="index-title-main"><i class="bi bi-list-columns-reverse me-2 text-primary"></i>REPORT INDEX & TABLE OF CONTENTS</h2>
                                    <div class="index-title-sub">Comprehensive Section Roadmap & Specific Information Locator</div>
                                </div>
                                <div>
                                    <span class="index-badge-tag"><i class="bi bi-compass me-1"></i>Page 2 Index</span>
                                </div>
                            </div>
                        </div>

                        <!-- TOC Table -->
                        <table class="index-table">
                            <thead>
                                <tr>
                                    <th style="width: 55px; text-align: center;">Sec #</th>
                                    <th>Report Section Title & Analytical Scope</th>
                                    <th style="width: 140px; text-align: center;">Category / Focus</th>
                                    <th style="width: 80px; text-align: center;">Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 1</td>
                                    <td>
                                        <span class="index-sec-title">1. Overall Institution Performance Summary</span>
                                        <div class="index-sec-desc">Executive summary of total enrolled, sat for exam, passed, failed, pass rate %, and GPA distribution.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-primary-subtle text-primary border border-primary-subtle index-tag-badge">Overview</span></td>
                                    <td class="text-center"><a href="#report-sec-1" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 2</td>
                                    <td>
                                        <span class="index-sec-title">2. Regular vs. Irregular Candidates Breakdown</span>
                                        <div class="index-sec-desc">Comparative analysis of fresh regular candidates vs. irregular/retake students & group matrix.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-info-subtle text-info-emphasis border border-info-subtle index-tag-badge">Comparative</span></td>
                                    <td class="text-center"><a href="#report-sec-2" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 3</td>
                                    <td>
                                        <span class="index-sec-title">3. Group Performance Comparison</span>
                                        <div class="index-sec-desc">Academic discipline breakdown across Science, Humanities, and Business Studies groups.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle index-tag-badge">Disciplines</span></td>
                                    <td class="text-center"><a href="#report-sec-3" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 4</td>
                                    <td>
                                        <span class="index-sec-title">4. Subject Wise Grade Analysis (Subsections 4.1 &ndash; 4.7+)</span>
                                        <div class="index-sec-desc">Detailed grade distribution (A+ through F), pass %, appeared, and passed counts by subject & stream.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-dark-subtle text-dark border border-dark-subtle index-tag-badge">Subject Grades</span></td>
                                    <td class="text-center"><a href="#report-sec-4" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 5</td>
                                    <td>
                                        <span class="index-sec-title">5. Student Merit List (Academic Ranking)</span>
                                        <div class="index-sec-desc">Examinees ranked in descending order of GPA (With 4th), Pure GPA (Without 4th), and grade totals.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-warning-subtle text-dark border border-warning-subtle index-tag-badge">Merit Ranking</span></td>
                                    <td class="text-center"><a href="#report-sec-5" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 6</td>
                                    <td>
                                        <span class="index-sec-title">6. Fail Analysis & Diagnostic Matrix (Without 4th Subject)</span>
                                        <div class="index-sec-desc">Official fail summary matrix and 4-column failure tables (All, Science, Humanities, Business Studies).</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-danger-subtle text-danger border border-danger-subtle index-tag-badge">Failure Matrix</span></td>
                                    <td class="text-center"><a href="#report-sec-6" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 7</td>
                                    <td>
                                        <span class="index-sec-title">7. One Subject Fail - Supplementary Eligible Students</span>
                                        <div class="index-sec-desc">Candidate roster of single-subject failures eligible for supplementary re-examination.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-warning-subtle text-danger border border-warning-subtle index-tag-badge">Supplementary</span></td>
                                    <td class="text-center"><a href="#report-sec-7" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Sec 8</td>
                                    <td>
                                        <span class="index-sec-title">8. Two Subject Fail Students Roster</span>
                                        <div class="index-sec-desc">Students failing in exactly two subjects with specific failed subject pairs identified.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-danger-subtle text-danger border border-danger-subtle index-tag-badge">2-Sub Remedial</span></td>
                                    <td class="text-center"><a href="#report-sec-8" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                ${(mFail.students || []).length > 0 ? `
                                <tr>
                                    <td class="text-center index-sec-num">Sec 9</td>
                                    <td>
                                        <span class="index-sec-title">9. Multi-Subject Fail Students Roster (3+ Subjects)</span>
                                        <div class="index-sec-desc">Examinees failing in 3 or more subjects requiring intensive academic intervention.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle index-tag-badge">Critical Intervention</span></td>
                                    <td class="text-center"><a href="#report-sec-9" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>` : ""}
                                <tr>
                                    <td class="text-center index-sec-num">Sec 10</td>
                                    <td>
                                        <span class="index-sec-title">10. Absent Candidates Analysis & Examination Roster</span>
                                        <div class="index-sec-desc">Complete roster of registered candidates who did not sit for the board examination.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-secondary-subtle text-secondary border border-secondary-subtle index-tag-badge">Absentee Audit</span></td>
                                    <td class="text-center"><a href="#report-sec-10" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                                <tr>
                                    <td class="text-center index-sec-num">Signoff</td>
                                    <td>
                                        <span class="index-sec-title">Institutional Verification & Official Signatures</span>
                                        <div class="index-sec-desc">Authentication and signature endorsements from Tabulator, Assistant, and Principal.</div>
                                    </td>
                                    <td class="text-center"><span class="badge bg-success-subtle text-success border border-success-subtle index-tag-badge">Certification</span></td>
                                    <td class="text-center"><a href="#report-sec-signoff" class="index-link-btn">Jump &rarr;</a></td>
                                </tr>
                            </tbody>
                        </table>

                        <!-- Quick Information Locator Box (Fast Finder for Specific Inquiries) -->
                        <div class="index-finder-box">
                            <div class="index-finder-title">
                                <i class="bi bi-search me-1"></i> Quick Information Locator &bull; How to Find Specific Info
                            </div>
                            <div class="index-finder-grid">
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Supplementary List</span>
                                    <span>Find single-subject fail students eligible for re-exam in <strong>Section 7</strong></span>
                                </div>
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Merit Ranks</span>
                                    <span>Find student academic positions, GPA 5 & rolls in <strong>Section 5</strong></span>
                                </div>
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Fail Diagnostics</span>
                                    <span>Identify high-failure subjects & groups in <strong>Section 6</strong></span>
                                </div>
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Subject Grades</span>
                                    <span>Find A+ counts & pass percentages per subject in <strong>Section 4</strong></span>
                                </div>
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Stream Totals</span>
                                    <span>Compare Science vs. Humanities vs. Biz outcomes in <strong>Section 3</strong></span>
                                </div>
                                <div class="index-finder-item">
                                    <span class="index-finder-badge">Absentee Rolls</span>
                                    <span>View roll numbers & parent names of absentees in <strong>Section 10</strong></span>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Index Page Footer -->
                    <div class="index-page-footer">
                        <div><i class="bi bi-info-circle me-1"></i>Click any section above to navigate directly on screen &bull; Save as PDF retains standard page flow</div>
                        <div class="fw-bold">Generated by iCT Center &bull; Contact: 01641711531 &bull; Page 2</div>
                    </div>
                </div>

                <!-- ========================================== -->
                <!-- PAGE 3+: MASTER REPORT LAYOUT & TABLES     -->
                <!-- ========================================== -->
                <table class="report-master-layout">
                    <thead class="report-master-thead">
                        <tr>
                            <th class="report-master-th">
                                <!-- Heading repeated on EVERY page in print -->
                                <div class="print-page-header">
                                    <h2 class="header-inst-name">${instDetails.instituteName}</h2>
                                    <div class="header-inst-meta">
                                        <span><i class="bi bi-geo-alt-fill text-danger me-1"></i><strong>Address:</strong> ${instDetails.address}</span>
                                        <span class="meta-dot">&bull;</span>
                                        <span><i class="bi bi-card-heading text-primary me-1"></i><strong>EIIN No:</strong> <span class="eiin-tag">${instDetails.eiin || "N/A"}</span></span>
                                        <span class="meta-dot">&bull;</span>
                                        <span><i class="bi bi-award-fill text-success me-1"></i><strong>Board:</strong> ${cleanBoard}</span>
                                    </div>
                                    <div class="header-line"></div>
                                </div>
                            </th>
                        </tr>
                    </thead>
                    <tbody class="report-master-tbody">
                        <tr>
                            <td class="report-master-td">
                                <div class="report-body-content">
                                    <!-- Document Header -->
                                    <div class="report-header d-flex justify-content-between align-items-center flex-wrap gap-2">
                    <div>
                        <div class="badge-pill-custom bg-dark text-white mb-1">
                            <i class="bi bi-award-fill text-warning me-1"></i> ACADEMIC EVALUATION & RESULT REPORT
                        </div>
                        <h3 class="fw-bold text-dark mb-0" style="letter-spacing: -0.5px;">INSTITUTION RESULT PERFORMANCE REPORT</h3>
                        <div class="text-muted small mt-1">
                            <strong>Exam:</strong> ${metaInfo.exam || "Examination"} ${metaInfo.year || ""} &bull; 
                            <strong>Board:</strong> ${(metaInfo.board || "Board").replace(/\s*\(Fast Server\)/i, "").trim()} &bull; 
                            <strong>Generated:</strong> ${new Date().toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' })} at ${new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                        </div>
                    </div>
                    <div class="text-end">
                        <span class="badge bg-primary-subtle text-primary border border-primary-subtle px-3 py-2 fs-7 fw-bold d-block mb-1">
                            Total Students: ${summary.total_students || 0}
                        </span>
                        <span class="badge bg-success-subtle text-success border border-success-subtle px-2 py-1 fs-8">
                            Pass Rate: ${summary.pass_rate || 0}%
                        </span>
                    </div>
                </div>

                <!-- Executive Visual KPI Dashboard -->
                <div class="row g-2 mb-3">
                    <div class="col-3">
                        <div class="kpi-card kpi-slate">
                            <div class="kpi-label">Total Enrolled</div>
                            <div class="kpi-num">${summary.total_students || 0}</div>
                            <div class="kpi-sub text-muted">Registered Candidates</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-blue">
                            <div class="kpi-label">Exam Attended</div>
                            <div class="kpi-num">${summary.appeared || 0}</div>
                            <div class="kpi-sub text-primary">Sat for Examination</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-green">
                            <div class="kpi-label">Total Passed</div>
                            <div class="kpi-num">${summary.passed || 0}</div>
                            <div class="kpi-sub text-success">${summary.pass_rate || 0}% of Attended</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-red">
                            <div class="kpi-label">Total Failed</div>
                            <div class="kpi-num">${summary.failed || 0}</div>
                            <div class="kpi-sub text-danger">${summary.fail_rate || 0}% of Attended</div>
                        </div>
                    </div>
                </div>

                <div class="row g-2 mb-3">
                    <div class="col-3">
                        <div class="kpi-card kpi-gold">
                            <div class="kpi-label">GPA 5.0 (With 4th)</div>
                            <div class="kpi-num">${gpa5WithCount}</div>
                            <div class="kpi-sub text-warning">${gpa5WithRate}% &bull; Official Board GPA</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-purple">
                            <div class="kpi-label">Golden A+ (Without 4th)</div>
                            <div class="kpi-num">${gpa5WithoutCount}</div>
                            <div class="kpi-sub" style="color: #6d28d9;">${gpa5WithoutRate}% &bull; Pure Main Subjects</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-orange">
                            <div class="kpi-label">1-Subject Fail</div>
                            <div class="kpi-num">${summary.one_subject_fail_count || 0}</div>
                            <div class="kpi-sub text-danger">Supplementary Eligible</div>
                        </div>
                    </div>
                    <div class="col-3">
                        <div class="kpi-card kpi-slate">
                            <div class="kpi-label">Absent Students</div>
                            <div class="kpi-num">${summary.absent || 0}</div>
                            <div class="kpi-sub text-secondary">${summary.absent_rate || 0}% Did Not Sit</div>
                        </div>
                    </div>
                </div>

                <!-- Plain-English Explanatory Guide Box for easy understanding -->
                <div class="guide-box">
                    <div class="fw-bold text-success mb-1" style="font-size: 11px;">
                        <i class="bi bi-info-circle-fill me-1"></i> How to Easily Understand These Metrics:
                    </div>
                    <div class="row g-2">
                        <div class="col-md-6">
                            <div class="guide-item">
                                &bull; <strong>GPA 5.0 (With 4th Subject):</strong> The official Board GPA including bonus points above 2.0 earned in the optional 4th subject.
                            </div>
                            <div class="guide-item">
                                &bull; <strong>Golden A+ (Without 4th Subject):</strong> Pure curriculum mastery where student achieved 5.00 GPA in compulsory & group subjects without needing optional bonus points.
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="guide-item">
                                &bull; <strong>4th Subject Safety Rule:</strong> Failing in the 4th optional subject does NOT fail a student who passed all main subjects (Official Board Rule).
                            </div>
                            <div class="guide-item">
                                &bull; <strong>One Subject Fail:</strong> Students who failed in only 1 subject and are eligible for single-subject supplementary re-examination.
                            </div>
                        </div>
                    </div>
                </div>

                <!-- SECTION 1: OVERALL INSTITUTION SUMMARY -->
                <div class="report-section" id="report-sec-1">
                    <div class="section-title">
                        <i class="bi bi-buildings-fill text-primary"></i> 1. Overall Institution Performance Summary
                    </div>
                    <div class="section-desc">Consolidated executive overview of all registered students, attendance, pass rates, and grade tiers.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th>Total Enrolled</th>
                                <th>Attended Exam</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="fw-bold">${summary.total_students || 0}</td>
                                <td class="fw-semibold text-primary">${summary.appeared || 0}</td>
                                <td class="fw-bold text-secondary">${summary.absent || 0}</td>
                                <td class="fw-bold text-success">${summary.passed || 0}</td>
                                <td class="fw-bold text-danger">${summary.failed || 0}</td>
                                <td class="fw-bold table-highlight-green">${summary.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${gpa5WithCount}</td>
                                <td class="fw-bold" style="color: #6d28d9; background-color: #f3e8ff;">${gpa5WithoutCount}</td>
                                <td class="fw-bold text-primary">${summary.one_subject_fail_count || 0}</td>
                                <td class="fw-bold text-danger">${summary.two_subject_fail_count || 0}</td>
                                <td class="text-muted">${summary.multi_subject_fail_count || 0}</td>
                            </tr>
                        </tbody>
                    </table>
                </div>

                <!-- SECTION 2: REGULAR VS IRREGULAR CANDIDATE PERFORMANCE -->
                <div class="report-section" id="report-sec-2">
                    <div class="section-title">
                        <i class="bi bi-people-fill text-primary"></i> 2. Regular vs. Irregular Candidates Breakdown
                    </div>
                    <div class="section-desc">Compares fresh regular batch candidates against candidates appearing for improvement or irregular examination.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th class="text-start ps-3">Candidate Category</th>
                                <th>Total Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td class="text-start ps-3 fw-bold text-primary">
                                    <span class="badge bg-primary-subtle text-primary border border-primary-subtle me-1">REGULAR</span> Fresh Batch
                                </td>
                                <td>${reg.total || 0}</td>
                                <td class="fw-semibold">${reg.appeared || 0}</td>
                                <td class="text-secondary fw-bold">${reg.absent || 0}</td>
                                <td class="fw-bold text-success">${reg.passed || 0}</td>
                                <td class="fw-bold text-danger">${reg.failed || 0}</td>
                                <td class="fw-bold text-success">${reg.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${reg.gpa5_count_with_4th ?? reg.gpa5_count ?? 0}</td>
                                <td class="fw-bold" style="color: #6d28d9;">${reg.gpa5_count_without_4th ?? 0}</td>
                                <td class="fw-semibold">${reg.one_fail_count || 0}</td>
                                <td class="text-danger fw-semibold">${reg.two_fail_count || 0}</td>
                                <td class="text-muted">${reg.multi_fail_count || 0}</td>
                            </tr>
                            <tr>
                                <td class="text-start ps-3 fw-bold text-warning">
                                    <span class="badge bg-warning-subtle text-dark border border-warning-subtle me-1">IRREGULAR</span> Improvement / Retake
                                </td>
                                <td>${irreg.total || 0}</td>
                                <td class="fw-semibold">${irreg.appeared || 0}</td>
                                <td class="text-secondary fw-bold">${irreg.absent || 0}</td>
                                <td class="fw-bold text-success">${irreg.passed || 0}</td>
                                <td class="fw-bold text-danger">${irreg.failed || 0}</td>
                                <td class="fw-bold text-success">${irreg.pass_rate || 0}%</td>
                                <td class="fw-bold table-highlight-gold">${irreg.gpa5_count_with_4th ?? irreg.gpa5_count ?? 0}</td>
                                <td class="fw-bold" style="color: #6d28d9;">${irreg.gpa5_count_without_4th ?? 0}</td>
                                <td class="fw-semibold">${irreg.one_fail_count || 0}</td>
                                <td class="text-danger fw-semibold">${irreg.two_fail_count || 0}</td>
                                <td class="text-muted">${irreg.multi_fail_count || 0}</td>
                            </tr>
                        </tbody>
                    </table>

                    ${groupType.length > 0 ? `
                    <div class="fw-bold text-secondary mb-1" style="font-size: 10px;">Group-Wise Regular vs. Irregular Detailed Matrix</div>
                    <table class="table table-bordered table-sm text-center mb-3">
                        <thead>
                            <tr>
                                <th class="text-start ps-2">Academic Group</th>
                                <th>Category</th>
                                <th>Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${groupType.map(r => `
                                <tr>
                                    <td class="text-start ps-2 fw-bold">${r.group}</td>
                                    <td><span class="badge ${r.type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${r.type}</span></td>
                                    <td>${r.total}</td>
                                    <td>${r.appeared}</td>
                                    <td class="text-secondary">${r.absent || 0}</td>
                                    <td class="text-success fw-bold">${r.passed}</td>
                                    <td class="text-danger fw-bold">${r.failed}</td>
                                    <td class="fw-bold">${r.pass_rate}%</td>
                                    <td class="table-highlight-gold">${r.gpa5_count_with_4th ?? r.gpa5_count ?? 0}</td>
                                    <td style="color: #6d28d9; font-weight: 700;">${r.gpa5_count_without_4th ?? 0}</td>
                                    <td>${r.one_fail_count}</td>
                                    <td>${r.two_fail_count}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : ""}
                </div>

                <!-- SECTION 3: GROUP PERFORMANCE COMPARISON -->
                <div class="report-section" id="report-sec-3">
                    <div class="section-title">
                        <i class="bi bi-grid-3x3-gap-fill text-primary"></i> 3. Group Performance Comparison
                    </div>
                    <div class="section-desc">Academic achievement comparison across Science, Humanities, and Business Studies groups.</div>
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th class="text-start ps-3">Academic Group</th>
                                <th>Total Enrolled</th>
                                <th>Attended</th>
                                <th>Absent</th>
                                <th>Passed</th>
                                <th>Failed</th>
                                <th>Pass Rate (%)</th>
                                <th>GPA 5.0 (With 4th)</th>
                                <th>Golden A+ (Without 4th)</th>
                                <th>1-Sub Fail</th>
                                <th>2-Sub Fail</th>
                                <th>3+ Sub Fail</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${Object.values(groupsSummary).map(grp => `
                                <tr>
                                    <td class="text-start ps-3 fw-bold text-dark">${grp.group}</td>
                                    <td>${grp.total}</td>
                                    <td class="fw-semibold">${grp.appeared}</td>
                                    <td class="text-secondary fw-bold">${grp.absent || 0}</td>
                                    <td class="text-success fw-bold">${grp.passed}</td>
                                    <td class="text-danger fw-bold">${grp.failed}</td>
                                    <td class="fw-bold text-success">${grp.pass_rate}%</td>
                                    <td class="fw-bold table-highlight-gold">${grp.gpa5_count_with_4th ?? grp.gpa5_count ?? 0}</td>
                                    <td class="fw-bold" style="color: #6d28d9;">${grp.gpa5_count_without_4th ?? 0}</td>
                                    <td>${grp.one_fail_count}</td>
                                    <td>${grp.two_fail_count}</td>
                                    <td class="text-muted">${grp.multi_fail_count}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>

                <!-- SECTION 4: SUBJECT WISE GRADE ANALYSIS WITH SEPARATE SECTIONS -->
                <div class="report-section" id="report-sec-4">
                    <div class="section-title">
                        <i class="bi bi-book-half text-primary"></i> 4. SUBJECT WISE GRADE ANALYSIS
                    </div>
                    <div class="section-desc">Comprehensive subject-level evaluation and grade distributions organized into dedicated sections for Candidate Categories (All, Regular, Irregular, Improvement) and Academic Disciplines (Business Studies, Humanities, Science).</div>

                    <!-- 4.1 All Candidates -->
                    ${renderReportSubjTable(
                        allCandidatesSubjs.length > 0 ? allCandidatesSubjs : subjList.filter(it => it.type === "ALL"),
                        "4.1 All Candidates (Overall All Subjects Combined)",
                        "ALL CANDIDATES",
                        "bg-dark text-white",
                        "Combined performance of all examinees across every subject in the institution."
                    )}

                    <!-- 4.2 Regular Candidates -->
                    ${renderReportSubjTable(
                        regularSubjs,
                        "4.2 Regular Candidates - Subject Performance",
                        "REGULAR",
                        "bg-primary text-white",
                        "Subject grade distribution for regular school/college candidates."
                    )}

                    <!-- 4.3 Irregular Candidates -->
                    ${renderReportSubjTable(
                        irregularSubjs,
                        "4.3 Irregular Candidates - Subject Performance",
                        "IRREGULAR",
                        "bg-warning text-dark",
                        "Subject grade breakdown and audit for irregular candidates sitting for board examinations."
                    )}

                    <!-- 4.4 Improvement Candidates -->
                    ${improvementSubjs.length > 0 ? renderReportSubjTable(
                        improvementSubjs,
                        "4.4 Improvement Candidates - Subject Performance",
                        "IMPROVEMENT",
                        "bg-info text-white",
                        "Subject performance for candidates sitting for GPA improvement."
                    ) : ""}

                    <!-- 4.5 Business Studies Stream -->
                    ${renderReportSubjTable(
                        bizStudiesSubjs,
                        "4.5 Business Studies Stream - Subject Performance",
                        "BUSINESS STUDIES",
                        "bg-success text-white",
                        "Detailed grade distribution for core and elective subjects in the Business Studies discipline."
                    )}

                    <!-- 4.6 Humanities Stream -->
                    ${renderReportSubjTable(
                        humanitiesSubjs,
                        "4.6 Humanities Stream - Subject Performance",
                        "HUMANITIES",
                        "bg-secondary text-white",
                        "Detailed grade distribution for core and elective subjects in the Humanities / Arts discipline."
                    )}

                    <!-- 4.7 Science Stream -->
                    ${renderReportSubjTable(
                        scienceSubjs,
                        "4.7 Science Stream - Subject Performance",
                        "SCIENCE",
                        "bg-primary text-white",
                        "Detailed grade distribution for core and elective subjects in the Science discipline."
                    )}

                    <!-- Other Streams (if any) -->
                    ${otherGroupsList.map((grpName, gIdx) => {
                        const grpSubjs = subjList.filter(it => it.group.toUpperCase() === grpName.toUpperCase() && it.type === "ALL");
                        return renderReportSubjTable(
                            grpSubjs,
                            `4.${8 + gIdx} ${grpName} Stream - Subject Performance`,
                            grpName,
                            "bg-dark text-white",
                            `Subject performance for ${grpName} group.`
                        );
                    }).join("")}
                </div>

                <!-- SECTION 5: STUDENT MERIT LIST -->
                <div class="report-section" id="report-sec-5">
                    <div class="section-title">
                        <i class="bi bi-trophy-fill text-warning"></i> 5. STUDENT MERIT LIST (ACADEMIC RANKING - TOTAL PASSED: ${(meritData.students || []).length})
                    </div>
                    <div class="section-desc">Students ranked in descending order of GPA (With 4th), Pure GPA (Without 4th), and grade distribution.</div>
                    ${(meritData.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm text-center">
                        <thead>
                            <tr>
                                <th style="width: 55px;">Rank</th>
                                <th style="width: 75px;">Roll No</th>
                                <th style="width: 85px;">Reg No</th>
                                <th class="text-start ps-2">Student Name</th>
                                <th style="width: 85px;">Group</th>
                                <th style="width: 75px;">Type</th>
                                <th style="width: 85px;">GPA (With 4th)</th>
                                <th style="width: 85px;">GPA (Pure)</th>
                                <th style="width: 100px;">Grades Brief</th>
                                <th class="text-start ps-2">Passed Subjects Summary</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(meritData.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold ${st.Merit_Position <= 3 ? 'table-highlight-gold' : ''}">${st.Merit_Position}</td>
                                    <td class="fw-bold">${st.Roll}</td>
                                    <td class="font-monospace text-muted">${st.Registration || "-"}</td>
                                    <td class="text-start ps-2 fw-semibold text-dark">${st.Name}</td>
                                    <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                                    <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="fw-bold ${st.Is_GPA5_With_4th ? 'table-highlight-gold' : 'text-primary'}">${st.GPA_With_4th}</td>
                                    <td class="fw-semibold ${st.Is_GPA5_Without_4th ? 'text-success fw-bold' : 'text-muted'}">${st.GPA_Without_4th}</td>
                                    <td class="small">${st.A_Plus_Count || 0} A+, ${st.A_Count || 0} A</td>
                                    <td class="text-start ps-2 small text-muted text-truncate" style="max-width: 220px;">${st.Passed_Grades_Summary || "-"}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No passed students to rank.</p>`}
                </div>

                <!-- SECTION 6: FAIL ANALYSIS (WITHOUT 4TH SUBJECT - ALL & GROUP-WISE) -->
                <div class="report-section" id="report-sec-6" style="page-break-inside: auto; break-inside: auto;">
                    <div class="section-title">
                        <i class="bi bi-pie-chart-fill text-danger"></i> 6. FAIL ANALYSIS (WITHOUT 4TH SUBJECT) - REGULAR & IRREGULAR BREAKDOWN
                    </div>
                    <div class="section-desc">Official 4-column failure tables showing failed subjects with Regular, Irregular, and Total student breakdown across All Candidates and Academic Groups (strictly excluding the 4th/optional subject).</div>

                    <!-- 6.1 One Subject Fail Tables -->
                    <div style="page-break-inside: avoid; break-inside: avoid; margin-bottom: 16px;">
                        <div style="background-color: #0f172a; color: #ffffff; padding: 5px 12px; border-radius: 4px 4px 0 0; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; display: flex; justify-content: space-between; align-items: center;">
                            <span>6.1 One Subject Fail (Without 4th Subject)</span>
                            <span style="background: #334155; padding: 2px 8px; border-radius: 3px; font-size: 8.5px;">All & Group Breakdown</span>
                        </div>
                        <div class="fail-main-grid" style="border: 1.5px solid #000; border-top: none; padding: 8px; background-color: #f8fafc; border-radius: 0 0 4px 4px;">
                            ${renderMainPrintFailTable("All", "One Subject Fail (Without 4th Subject)", failTables.one_subject_fail?.all || {})}
                            ${availableGroups.map(grp => renderMainPrintFailTable(`Group: ${grp}`, "One Subject Fail (Without 4th Subject)", failTables.one_subject_fail?.by_group?.[grp] || {})).join("")}
                        </div>
                    </div>

                    <!-- 6.2 Two Subject Fail Tables -->
                    <div style="page-break-inside: avoid; break-inside: avoid; margin-bottom: 16px;">
                        <div style="background-color: #0f172a; color: #ffffff; padding: 5px 12px; border-radius: 4px 4px 0 0; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; display: flex; justify-content: space-between; align-items: center;">
                            <span>6.2 Two Subject Fail (Without 4th Subject)</span>
                            <span style="background: #334155; padding: 2px 8px; border-radius: 3px; font-size: 8.5px;">All & Group Breakdown</span>
                        </div>
                        <div class="fail-main-grid" style="border: 1.5px solid #000; border-top: none; padding: 8px; background-color: #f8fafc; border-radius: 0 0 4px 4px;">
                            ${renderMainPrintFailTable("All", "Two Subject Fail (Without 4th Subject)", failTables.two_subject_fail?.all || {})}
                            ${availableGroups.map(grp => renderMainPrintFailTable(`Group: ${grp}`, "Two Subject Fail (Without 4th Subject)", failTables.two_subject_fail?.by_group?.[grp] || {})).join("")}
                        </div>
                    </div>

                    <!-- 6.3 Three or More Subject Fail Tables -->
                    <div style="page-break-inside: avoid; break-inside: avoid; margin-bottom: 16px;">
                        <div style="background-color: #0f172a; color: #ffffff; padding: 5px 12px; border-radius: 4px 4px 0 0; font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: 0.3px; display: flex; justify-content: space-between; align-items: center;">
                            <span>6.3 Three or More Subject Fail (Without 4th Subject)</span>
                            <span style="background: #334155; padding: 2px 8px; border-radius: 3px; font-size: 8.5px;">All & Group Breakdown</span>
                        </div>
                        <div class="fail-main-grid" style="border: 1.5px solid #000; border-top: none; padding: 8px; background-color: #f8fafc; border-radius: 0 0 4px 4px;">
                            ${renderMainPrintFailTable("All", "Three or More Subject Fail (Without 4th Subject)", failTables.three_plus_subject_fail?.all || {})}
                            ${availableGroups.map(grp => renderMainPrintFailTable(`Group: ${grp}`, "Three or More Subject Fail (Without 4th Subject)", failTables.three_plus_subject_fail?.by_group?.[grp] || {})).join("")}
                        </div>
                    </div>
                </div>

                <!-- SECTION 7: ONE SUBJECT FAIL STUDENTS -->
                <div class="report-section" id="report-sec-7">
                    <div class="section-title">
                        <i class="bi bi-exclamation-octagon-fill text-warning"></i> 7. One Subject Fail - Supplementary Eligible Students (Total: ${oneFail.total || 0})
                    </div>
                    <div class="section-desc">These students failed in only ONE subject. Remedial coaching and supplementary examination will allow them to pass easily.</div>
                    ${(oneFail.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 140px;">Failed Subject</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(oneFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="fw-bold text-danger table-highlight-red text-center">${st.Failed_Subject}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No students failed in only one subject.</p>`}
                </div>

                <!-- SECTION 8: TWO SUBJECT FAIL STUDENTS -->
                <div class="report-section" id="report-sec-8">
                    <div class="section-title">
                        <i class="bi bi-exclamation-triangle-fill text-danger"></i> 8. Two Subject Fail Students (Total: ${twoFail.total || 0})
                    </div>
                    <div class="section-desc">Students who failed in exactly two subjects and require intervention in these specific subjects.</div>
                    ${(twoFail.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 180px;">Failed Subjects</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(twoFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="fw-bold text-danger table-highlight-red text-center">${(st.Failed_Subjects || []).join(" & ")}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No students failed in two subjects.</p>`}
                </div>

                <!-- SECTION 9: 3+ SUBJECT FAIL STUDENTS (IF ANY) -->
                ${(mFail.students || []).length > 0 ? `
                <div class="report-section" id="report-sec-9">
                    <div class="section-title">
                        <i class="bi bi-x-circle-fill text-secondary"></i> 9. Multi-Subject Fail Students (3+ Subjects) - Total: ${mFail.total || 0}
                    </div>
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th>Student Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th style="width: 80px;">Fails Count</th>
                                <th>Failed Subjects</th>
                                <th>Passed Subject Grades</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(mFail.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll}</td>
                                    <td class="fw-semibold">${st.Name}</td>
                                    <td class="text-center">${st.Group}</td>
                                    <td class="text-center">${st.Type || "REGULAR"}</td>
                                    <td class="fw-bold text-danger text-center">${st.Failed_Count}</td>
                                    <td class="text-danger fw-semibold">${(st.Failed_Subjects || []).join(", ")}</td>
                                    <td class="small text-muted">${st.Passed_Grades_Summary}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>` : ""}

                <!-- SECTION 10: ABSENT CANDIDATES ANALYSIS & ROSTER -->
                <div class="report-section" id="report-sec-10">
                    <div class="section-title">
                        <i class="bi bi-person-x-fill text-secondary"></i> 10. Absent Candidates Roster (Did Not Sit for Exam) - Total: ${absentData.total_absent || summary.absent || 0}
                    </div>
                    <div class="section-desc">Registered candidates who were absent during the examination.</div>
                    
                    <div class="p-2 mb-3 bg-light border rounded d-flex justify-content-between align-items-center flex-wrap gap-2" style="font-size: 10.5px;">
                        <div>
                            <strong>Total Absent:</strong> <span class="badge bg-secondary">${absentData.total_absent || summary.absent || 0}</span> (${absentData.absent_rate || summary.absent_rate || 0}% of Enrolled) &bull; 
                            <strong>Regular Absent:</strong> ${absentData.regular_absent || 0} &bull; 
                            <strong>Irregular Absent:</strong> ${absentData.irregular_absent || 0}
                        </div>
                        <div>
                            <strong>Total Appeared (Attended Exam):</strong> <span class="badge bg-success">${absentData.appeared || summary.appeared || 0}</span>
                        </div>
                    </div>

                    ${(absentData.students || []).length > 0 ? `
                    <table class="table table-bordered table-sm">
                        <thead>
                            <tr>
                                <th style="width: 80px;">Roll No</th>
                                <th style="width: 90px;">Reg No</th>
                                <th>Student Name</th>
                                <th>Father's Name</th>
                                <th style="width: 90px;">Group</th>
                                <th style="width: 80px;">Type</th>
                                <th>Institute</th>
                                <th style="width: 75px;">Status</th>
                                <th>Remarks</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${(absentData.students || []).map(st => `
                                <tr>
                                    <td class="fw-bold text-center">${st.Roll || "-"}</td>
                                    <td class="text-center font-monospace">${st.Registration || "-"}</td>
                                    <td class="fw-semibold text-dark">${st.Name || "N/A"}</td>
                                    <td class="text-muted small">${st["Father Name"] || st.Father_Name || "-"}</td>
                                    <td class="text-center">${st.Group || "GENERAL"}</td>
                                    <td class="text-center"><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border border-primary-subtle' : 'bg-warning-subtle text-dark border border-warning-subtle'}">${st.Type || "REGULAR"}</span></td>
                                    <td class="small text-truncate" style="max-width: 140px;">${st.Institute || "-"}</td>
                                    <td class="fw-bold text-center text-secondary"><span class="badge bg-secondary">ABSENT</span></td>
                                    <td class="small text-muted">${st.Remarks || "Did not sit for examination"}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>` : `<p class="text-muted fst-italic">No absent candidates recorded in this examination result.</p>`}
                </div>

                <!-- Footer Signoff Section -->
                <div class="signoff-section" id="report-sec-signoff">
                    <div class="row pt-4 text-center text-muted" style="font-size: 8.5pt;">
                        <div class="col-4">
                            <div class="sign-line">Prepared by (Staff)</div>
                        </div>
                        <div class="col-4">
                            <div class="sign-line">Checked by (Tabulator)</div>
                        </div>
                        <div class="col-4">
                            <div class="sign-line">Principal / Head of Institution</div>
                        </div>
                    </div>
                </div>
                                </div>
                            </td>
                        </tr>
                    </tbody>
                    <tfoot class="report-master-tfoot">
                        <tr>
                            <th class="report-master-tf">
                                <!-- Footer repeated on EVERY page in print -->
                                <div class="print-page-footer">
                                    <div class="footer-line"></div>
                                    <div class="footer-row">
                                        <div class="footer-brand">
                                            <i class="bi bi-shield-check me-1"></i>Genrated By iCT Center &nbsp;&bull;&nbsp; Contact: 01641711531
                                        </div>
                                        <div class="footer-meta">
                                            <span>Comprehensive Academic Performance Report &bull; EIIN: ${instDetails.eiin || "N/A"} &bull; A4 Document</span>
                                        </div>
                                    </div>
                                </div>
                            </th>
                        </tr>
                    </tfoot>
                </table>

                <!-- Bottom Print Button (Hidden on paper) -->
                <div class="no-print text-center my-4 pt-3 border-top">
                    <button class="btn btn-lg px-5 shadow-sm fw-bold text-white" style="background-color: #1e3a5f;" onclick="window.print()">
                        <i class="bi bi-printer-fill me-2"></i> Print / Save as PDF
                    </button>
                    <button class="btn btn-outline-secondary btn-lg px-4 ms-2" onclick="window.close()">
                        Close Window
                    </button>
                </div>
            </div>
        </body>
        </html>
        `;
        printWin.document.write(html);
        printWin.document.close();
    }

    // ==========================================
    // SEPARATE PDF DOWNLOAD FUNCTIONALITY FOR EVERY OPTION
    // ==========================================

    function setupSeparateOptionPdfButtons(analysisData, metaInfo) {
        const optionButtons = [
            { id: "pdfBtnSubjAll", opt: "subj_all" },
            { id: "pdfBtnSubjReg", opt: "subj_reg" },
            { id: "pdfBtnSubjIrreg", opt: "subj_irreg" },
            { id: "pdfBtnMerit", opt: "merit" },
            { id: "pdfBtn1Fail", opt: "1fail" },
            { id: "pdfBtn2Fail", opt: "2fail" },
            { id: "pdfBtnMFail", opt: "mfail" },
            { id: "pdfBtnAbsent", opt: "absent" },
            { id: "pdfBtnSummary", opt: "summary" },
            { id: "pdfBtnRegIrreg", opt: "reg_irreg" }
        ];

        optionButtons.forEach(b => {
            const btnEl = document.getElementById(b.id);
            if (btnEl) {
                btnEl.onclick = (e) => {
                    e.preventDefault();
                    downloadOptionPdf(b.opt, analysisData, metaInfo);
                };
            }
        });
    }

    // Generic PDF Report Window Generator (A4, print-optimized with instant Save as PDF)
    function generatePdfReportWindow(title, subtitle, metaInfo, kpiCards, contentHtml, noteText = "") {
        const printWin = window.open("", "_blank");
        const instDetails = getInstitutionHeaderDetails(latestAnalysisData, metaInfo);
        if (!printWin) {
            alert("Please allow popups in your browser to open and save the PDF document.");
            return;
        }

        const cleanBoard = (metaInfo && metaInfo.board ? metaInfo.board : "Board").replace(/\s*\(Fast Server\)/i, "").trim();
        const examLabel = [metaInfo && metaInfo.exam ? metaInfo.exam : "Analyzed File", metaInfo && metaInfo.year ? metaInfo.year : ""].filter(Boolean).join(" ").trim();
        const cleanInstName = (instDetails.instituteName || "Institution").replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
        const eiinPrefix = instDetails.eiin ? `${instDetails.eiin}, ` : "";
        const cleanSectionTitle = (title || "Report").replace(/[:*?"<>|\/\\]/g, "").replace(/\s+/g, " ").trim();
        const safeDocTitle = `${eiinPrefix}${cleanInstName} ${cleanSectionTitle} - ${examLabel}`;
        const nowStr = new Date().toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' });

        const html = `
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>${safeDocTitle}</title>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
            <style>
                @page {
                    size: A4 portrait;
                    margin: 10mm 9mm 12mm 9mm;
                }
                * {
                    box-sizing: border-box;
                }
                body {
                    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                    font-size: 9pt;
                    line-height: 1.35;
                    color: #1e293b;
                    background-color: #f1f5f9;
                    padding: 18px 14px;
                    margin: 0;
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                .report-container {
                    max-width: 1080px;
                    margin: 0 auto;
                    background: #ffffff;
                    padding: 22px 26px;
                    border-radius: 12px;
                    box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.08);
                    border: 1px solid #e2e8f0;
                }
                table.report-master-layout {
                    width: 100% !important;
                    border-collapse: collapse !important;
                    border: none !important;
                    margin: 0 !important;
                    padding: 0 !important;
                    background: transparent !important;
                }
                table.report-master-layout > thead {
                    display: table-header-group !important;
                }
                table.report-master-layout > tfoot {
                    display: table-footer-group !important;
                }
                table.report-master-layout > thead > tr > th,
                table.report-master-layout > thead > tr > td,
                table.report-master-layout > tbody > tr > td,
                table.report-master-layout > tfoot > tr > th,
                table.report-master-layout > tfoot > tr > td {
                    border: none !important;
                    padding: 0 !important;
                    margin: 0 !important;
                    background: transparent !important;
                    font-weight: normal;
                    text-align: left;
                }
                .print-page-header {
                    width: 100%;
                    padding-bottom: 7px;
                    margin-bottom: 11px;
                    text-align: center;
                }
                .header-inst-name {
                    font-size: 16.5pt;
                    font-weight: 800;
                    color: #0f2942;
                    margin: 0 0 3px 0;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    line-height: 1.2;
                }
                .header-inst-meta {
                    font-size: 8.5pt;
                    color: #334155;
                    font-weight: 600;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    gap: 12px;
                    flex-wrap: wrap;
                    margin-top: 3px;
                }
                .meta-dot {
                    color: #94a3b8;
                    font-weight: bold;
                }
                .eiin-tag {
                    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
                    font-size: 9pt;
                    font-weight: 800;
                    background: #e2e8f0;
                    color: #0f2942;
                    padding: 1px 7px;
                    border-radius: 4px;
                    letter-spacing: 0.5px;
                    border: 1px solid #cbd5e1;
                    display: inline-block;
                }
                .header-line {
                    height: 2.5px;
                    background: linear-gradient(90deg, #0f2942 0%, #1e3a5f 60%, #3b82f6 100%);
                    margin-top: 8px;
                    border-radius: 2px;
                }
                .print-page-footer {
                    width: 100%;
                    padding-top: 6px;
                    margin-top: 10px;
                }
                .footer-line {
                    height: 1.5px;
                    background: #cbd5e1;
                    margin-bottom: 5px;
                }
                .footer-row {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                    font-size: 8.2pt;
                    color: #475569;
                }
                .footer-brand {
                    font-weight: 700;
                    color: #0f2942;
                    letter-spacing: 0.2px;
                }
                .footer-meta {
                    color: #64748b;
                    font-size: 8pt;
                }
                .report-header {
                    border-bottom: 1.5px solid #cbd5e1;
                    padding-bottom: 10px;
                    margin-bottom: 14px;
                }
                .report-note-box {
                    background-color: #f8fafc;
                    border: 1px solid #e2e8f0;
                    border-left: 3.5px solid #3b82f6;
                    border-radius: 6px;
                    padding: 8px 12px;
                    font-size: 8.5pt;
                    color: #475569;
                    margin-bottom: 14px;
                }
                .kpi-grid {
                    display: flex;
                    gap: 10px;
                    margin-bottom: 14px;
                    flex-wrap: wrap;
                }
                .kpi-card {
                    flex: 1 1 130px;
                    background: #ffffff;
                    border: 1px solid #e2e8f0;
                    border-top: 3px solid #2563eb;
                    border-radius: 8px;
                    padding: 8px 10px;
                    text-align: center;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
                }
                .kpi-card.kpi-success { border-top-color: #16a34a; }
                .kpi-card.kpi-danger { border-top-color: #dc2626; }
                .kpi-card.kpi-warning { border-top-color: #d97706; }
                .kpi-card.kpi-dark { border-top-color: #0f2942; }
                .kpi-card .kpi-num {
                    font-size: 16pt;
                    font-weight: 800;
                    line-height: 1.2;
                    margin-top: 2px;
                    margin-bottom: 1px;
                    color: #0f2942;
                }
                .kpi-card .kpi-label {
                    font-size: 7.5pt;
                    font-weight: 700;
                    text-transform: uppercase;
                    letter-spacing: 0.5px;
                    color: #64748b;
                }
                table {
                    width: 100%;
                    margin-bottom: 14px;
                    font-size: 8.5pt;
                    border-collapse: collapse;
                }
                th {
                    background-color: #1e3a5f !important;
                    color: #ffffff !important;
                    font-weight: 600;
                    text-transform: uppercase;
                    font-size: 7.8pt;
                    letter-spacing: 0.3px;
                    padding: 5px 6px;
                    border: 1px solid #334155 !important;
                    text-align: center;
                    vertical-align: middle;
                }
                td {
                    padding: 4.5px 6px;
                    border: 1px solid #cbd5e1;
                    text-align: center;
                    vertical-align: middle;
                    color: #1e293b;
                }
                tbody tr:nth-child(even) { background-color: #f8fafc; }
                tbody tr:hover { background-color: #f1f5f9; }
                .badge-gpa-gold { background-color: #fef3c7 !important; color: #92400e !important; font-weight: 700; border: 1px solid #fde68a !important; border-radius: 3px; padding: 1px 4px; }
                .badge-gpa-green { background-color: #dcfce7 !important; color: #166534 !important; font-weight: 700; border: 1px solid #bbf7d0 !important; border-radius: 3px; padding: 1px 4px; }
                .badge-gpa-blue { background-color: #e0f2fe !important; color: #0369a1 !important; font-weight: 600; border: 1px solid #bae6fd !important; border-radius: 3px; padding: 1px 4px; }
                .badge-gpa-red { background-color: #fee2e2 !important; color: #b91c1c !important; font-weight: 700; border: 1px solid #fecaca !important; border-radius: 3px; padding: 1px 4px; }
                .signoff-section {
                    margin-top: 30px;
                    padding-top: 15px;
                    border-top: 1px dashed #cbd5e1;
                    font-size: 8.5pt;
                    color: #475569;
                    page-break-inside: avoid;
                    break-inside: avoid;
                }
                .sign-line {
                    border-top: 1.5px solid #0f2942;
                    margin: 0 15px;
                    padding-top: 4px;
                    font-weight: 700;
                    color: #0f2942;
                }
                @media print {
                    .no-print { display: none !important; }
                    body { padding: 0 !important; margin: 0 !important; background: #ffffff !important; }
                    .report-container { max-width: 100% !important; padding: 0 !important; margin: 0 !important; box-shadow: none !important; border: none !important; border-radius: 0 !important; }
                    tr { page-break-inside: avoid !important; break-inside: avoid !important; }
                    .kpi-card, .kpi-grid, .fail-main-card, .report-subsection, .signoff-section { page-break-inside: avoid !important; break-inside: avoid !important; }
                }
            </style>
        </head>
        <body>
            <div class="report-container">
                <!-- Action Toolbar (Hidden in print) -->
                <div class="no-print d-flex justify-content-between align-items-center p-3 mb-3 bg-white border border-slate-200 rounded-3 shadow-sm">
                    <div class="d-flex align-items-center gap-2">
                        <span class="badge px-3 py-2 text-white" style="background-color: #1e3a5f; font-size: 11px;">
                            <i class="bi bi-file-earmark-pdf-fill me-1"></i> A4 PDF Ready
                        </span>
                        <span class="text-secondary small">Click <strong>Print / Save as PDF</strong> below to download or print this document on A4.</span>
                    </div>
                    <div class="d-flex gap-2">
                        <button class="btn btn-sm px-4 fw-bold text-white shadow-sm" style="background-color: #1e3a5f;" onclick="window.print()">
                            <i class="bi bi-printer-fill me-1"></i> Print / Save as PDF
                        </button>
                        <button class="btn btn-outline-secondary btn-sm px-3" onclick="window.close()">
                            Close
                        </button>
                    </div>
                </div>

                <table class="report-master-layout">
                    <thead class="report-master-thead">
                        <tr>
                            <th class="report-master-th">
                                <!-- Heading repeated on EVERY page in print -->
                                <div class="print-page-header">
                                    <h2 class="header-inst-name">${instDetails.instituteName}</h2>
                                    <div class="header-inst-meta">
                                        <span><i class="bi bi-geo-alt-fill text-danger me-1"></i><strong>Address:</strong> ${instDetails.address}</span>
                                        <span class="meta-dot">&bull;</span>
                                        <span><i class="bi bi-card-heading text-primary me-1"></i><strong>EIIN No:</strong> <span class="eiin-tag">${instDetails.eiin || "N/A"}</span></span>
                                        <span class="meta-dot">&bull;</span>
                                        <span><i class="bi bi-award-fill text-success me-1"></i><strong>Board:</strong> ${cleanBoard}</span>
                                    </div>
                                    <div class="header-line"></div>
                                </div>
                            </th>
                        </tr>
                    </thead>
                    <tbody class="report-master-tbody">
                        <tr>
                            <td class="report-master-td">
                                <div class="report-body-content">
                                    <!-- Header -->
                                    <div class="report-header">
                                        <div class="d-flex justify-content-between align-items-center">
                                            <div>
                                                <h4 class="fw-bold mb-0 text-dark" style="letter-spacing: -0.3px;">${title}</h4>
                                                <div class="text-primary fw-semibold mt-1" style="font-size: 11px;">
                                                    ${metaInfo.exam || "Examination"} ${metaInfo.year || ""} &bull; ${cleanBoard}
                                                </div>
                                                ${subtitle ? `<div class="text-muted small mt-0.5">${subtitle}</div>` : ""}
                                            </div>
                                            <div class="text-end text-secondary small" style="font-size: 8.5pt; line-height: 1.45;">
                                                <div><strong>Report Date:</strong> ${nowStr}</div>
                                                <div><strong>Academic Session:</strong> ${metaInfo.year || new Date().getFullYear()}</div>
                                                <div class="text-dark fw-bold mt-0.5">Academic Performance Record</div>
                                            </div>
                                        </div>
                                    </div>

                                    ${kpiCards ? `<div class="kpi-grid">${kpiCards}</div>` : ""}

                                    ${noteText ? `<div class="report-note-box">${noteText}</div>` : ""}

                                    <!-- Content Data Table -->
                                    <div class="table-responsive">
                                        ${contentHtml}
                                    </div>

                                    <!-- Institutional Signoff -->
                                    <div class="signoff-section">
                                        <div class="row text-center pt-3">
                                            <div class="col-4">
                                                <div class="sign-line">Prepared by (Assistant)</div>
                                            </div>
                                            <div class="col-4">
                                                <div class="sign-line">Checked by (Tabulator)</div>
                                            </div>
                                            <div class="col-4">
                                                <div class="sign-line">Head of Institution / Principal</div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </td>
                        </tr>
                    </tbody>
                    <tfoot class="report-master-tfoot">
                        <tr>
                            <th class="report-master-tf">
                                <!-- Footer repeated on EVERY page in print -->
                                <div class="print-page-footer">
                                    <div class="footer-line"></div>
                                    <div class="footer-row">
                                        <div class="footer-brand">
                                            <i class="bi bi-shield-check me-1"></i>Genrated By iCT Center &nbsp;&bull;&nbsp; Contact: 01641711531
                                        </div>
                                        <div class="footer-meta">
                                            <span>${cleanSectionTitle} &bull; EIIN: ${instDetails.eiin || "N/A"} &bull; Page Print Copy</span>
                                        </div>
                                    </div>
                                </div>
                            </th>
                        </tr>
                    </tfoot>
                </table>

                <!-- Bottom Print Button (Hidden on paper) -->
                <div class="no-print text-center my-4 pt-3 border-top">
                    <button class="btn btn-lg px-5 shadow-sm fw-bold text-white" style="background-color: #1e3a5f;" onclick="window.print()">
                        <i class="bi bi-printer-fill me-1"></i> Print / Save as PDF
                    </button>
                    <button class="btn btn-outline-secondary btn-lg px-4 ms-2" onclick="window.close()">
                        Close Window
                    </button>
                </div>
            </div>

            <script>
                window.addEventListener('DOMContentLoaded', () => {
                    setTimeout(() => {
                        window.focus();
                        window.print();
                    }, 400);
                });
            </script>
        </body>
        </html>
        `;

        printWin.document.write(html);
        printWin.document.close();
    }

    // Master dispatcher for exporting each separate option as a clean PDF
    function downloadOptionPdf(optionKey, analysisData, metaInfo) {
        const data = analysisData || latestAnalysisData;
        if (!data) {
            alert("No result analysis data is available to export. Please extract or upload results first.");
            return;
        }

        const meta = metaInfo || {
            exam: (examSelect && examSelect.options[examSelect.selectedIndex]?.text) || "HSC",
            year: (yearSelect && yearSelect.value) || "2026",
            board: (boardSelect && boardSelect.options[boardSelect.selectedIndex]?.text) || "Board"
        };

        if (optionKey === "subj_all" || optionKey === "subj_reg" || optionKey === "subj_irreg") {
            const prefix = optionKey === "subj_all" ? "all" : (optionKey === "subj_reg" ? "reg" : "irreg");
            const targetType = optionKey === "subj_all" ? "ALL" : (optionKey === "subj_reg" ? "REGULAR" : "IRREGULAR");
            const typeLabel = optionKey === "subj_all" ? "Subject Wise Grade Analysis" : (optionKey === "subj_reg" ? "Regular Candidates Subject Analysis" : "Irregular Candidates Subject Analysis");

            const groupFilter = document.getElementById(`${prefix}SubjGroupFilter`);
            const searchInput = document.getElementById(`${prefix}SubjSearch`);
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();
            const allSubjs = data.subject_analysis || [];

            const renderSubjSectionTable = (items, sectionHeading, badgeLabel, badgeClass, descText = "") => {
                if (!items || items.length === 0) return "";
                const totApp = items.reduce((a, b) => a + (b.appeared || 0), 0);
                const totPass = items.reduce((a, b) => a + (b.passed || 0), 0);
                const totFail = items.reduce((a, b) => a + (b.failed || 0), 0);
                const overallRate = totApp > 0 ? ((totPass / totApp) * 100).toFixed(2) : "0.00";

                return `
                    <div class="mb-4">
                        <div class="d-flex justify-content-between align-items-center mb-1 pb-1 border-bottom">
                            <div>
                                <span class="badge ${badgeClass} me-2 font-monospace" style="font-size: 9px;">${badgeLabel}</span>
                                <strong class="text-dark" style="font-size: 11.5px;">${sectionHeading}</strong>
                            </div>
                            <div class="small text-muted font-monospace">
                                Subjects: <strong>${items.length}</strong> | Appeared: <strong>${totApp}</strong> | Passed: <strong class="text-success">${totPass}</strong> | Failed: <strong class="text-danger">${totFail}</strong> | Pass Rate: <strong class="text-primary">${overallRate}%</strong>
                            </div>
                        </div>
                        ${descText ? `<div class="small text-muted mb-2 fst-italic" style="font-size: 9px;">${descText}</div>` : ""}
                        <table class="table table-bordered table-sm mb-3">
                            <thead>
                                <tr>
                                    <th class="text-start ps-2">Subject Code & Name</th>
                                    <th>Group</th>
                                    <th>Type</th>
                                    <th>Appeared</th>
                                    <th>Passed</th>
                                    <th>Failed</th>
                                    <th>Pass Rate</th>
                                    <th class="badge-gpa-gold">A+</th>
                                    <th class="badge-gpa-green">A</th>
                                    <th class="badge-gpa-green">A-</th>
                                    <th class="badge-gpa-blue">B</th>
                                    <th class="badge-gpa-blue">C</th>
                                    <th>D</th>
                                    <th class="badge-gpa-red">F</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${items.map(it => `
                                    <tr>
                                        <td class="text-start ps-2 fw-semibold text-dark">${it.subject}</td>
                                        <td><span class="badge bg-light text-secondary border">${it.group}</span></td>
                                        <td><span class="badge bg-light text-dark border">${it.type}</span></td>
                                        <td class="fw-semibold">${it.appeared}</td>
                                        <td class="text-success fw-bold">${it.passed}</td>
                                        <td class="${it.failed > 0 ? 'text-danger fw-bold' : 'text-muted'}">${it.failed}</td>
                                        <td class="fw-bold">${it.pass_rate}%</td>
                                        <td class="badge-gpa-gold">${(it.grades || {})["A+"] || 0}</td>
                                        <td class="badge-gpa-green">${(it.grades || {})["A"] || 0}</td>
                                        <td class="badge-gpa-green">${(it.grades || {})["A-"] || 0}</td>
                                        <td class="badge-gpa-blue">${(it.grades || {})["B"] || 0}</td>
                                        <td class="badge-gpa-blue">${(it.grades || {})["C"] || 0}</td>
                                        <td>${(it.grades || {})["D"] || 0}</td>
                                        <td class="${(it.grades || {})['F'] > 0 ? 'text-danger fw-bold' : 'text-muted'}">${(it.grades || {})["F"] || 0}</td>
                                    </tr>
                                `).join("")}
                            </tbody>
                        </table>
                    </div>
                `;
            };

            let tableHtml = "";
            let kpiHtml = "";

            if (optionKey === "subj_all" && selectedGroup === "ALL" && !query) {
                // Separate different sections for "All", "IMPROVEMENT", "IRREGULAR", "REGULAR", "BUSINESS STUDIES", "HUMANITIES", "SCIENCE"
                const allCand = allSubjs.filter(it => it.group === "ALL" && it.type === "ALL");
                const allCandList = allCand.length > 0 ? allCand : allSubjs.filter(it => it.type === "ALL");
                const regularList = allSubjs.filter(it => it.group === "ALL" && it.type === "REGULAR");
                const irregularList = allSubjs.filter(it => it.group === "ALL" && it.type === "IRREGULAR");
                const improvementList = allSubjs.filter(it => it.group === "ALL" && it.type === "IMPROVEMENT");
                const bizList = allSubjs.filter(it => it.group.toUpperCase() === "BUSINESS STUDIES" && it.type === "ALL");
                const humList = allSubjs.filter(it => it.group.toUpperCase() === "HUMANITIES" && it.type === "ALL");
                const sciList = allSubjs.filter(it => it.group.toUpperCase() === "SCIENCE" && it.type === "ALL");

                const standardGrp = ["ALL", "BUSINESS STUDIES", "HUMANITIES", "SCIENCE"];
                const otherGroups = [...new Set(allSubjs.map(it => it.group))].filter(g => g && !standardGrp.includes(g.toUpperCase()));

                const totApp = allCandList.reduce((a, b) => a + (b.appeared || 0), 0);
                const totPass = allCandList.reduce((a, b) => a + (b.passed || 0), 0);
                const totFail = allCandList.reduce((a, b) => a + (b.failed || 0), 0);
                const passRate = totApp > 0 ? ((totPass / totApp) * 100).toFixed(2) : "0.00";

                kpiHtml = `
                    <div class="kpi-card"><div class="kpi-label">Total Subjects</div><div class="kpi-num text-dark">${allCandList.length}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Total Appeared</div><div class="kpi-num text-primary">${totApp}</div></div>
                    <div class="kpi-card kpi-success"><div class="kpi-label">Total Passed</div><div class="kpi-num text-success">${totPass}</div></div>
                    <div class="kpi-card kpi-danger"><div class="kpi-label">Total Failed</div><div class="kpi-num text-danger">${totFail}</div></div>
                    <div class="kpi-card kpi-warning"><div class="kpi-label">Pass Rate</div><div class="kpi-num text-warning">${passRate}%</div></div>
                `;

                tableHtml = `
                    ${renderSubjSectionTable(allCandList, "1. All Candidates (Overall All Subjects Combined)", "ALL CANDIDATES", "bg-dark text-white", "Combined evaluation of all candidates across all subjects.")}
                    ${renderSubjSectionTable(regularList, "2. Regular Candidates - Subject Performance", "REGULAR", "bg-primary text-white", "Subject grade distribution for regular candidates.")}
                    ${renderSubjSectionTable(irregularList, "3. Irregular Candidates - Subject Performance", "IRREGULAR", "bg-warning text-dark", "Subject grade distribution for irregular examinees.")}
                    ${improvementList.length > 0 ? renderSubjSectionTable(improvementList, "4. Improvement Candidates - Subject Performance", "IMPROVEMENT", "bg-info text-white", "Subject performance for candidates sitting for GPA improvement.") : ""}
                    ${renderSubjSectionTable(bizList, "5. Business Studies Stream - Subject Performance", "BUSINESS STUDIES", "bg-success text-white", "Detailed grade distribution for core and elective subjects in Business Studies.")}
                    ${renderSubjSectionTable(humList, "6. Humanities Stream - Subject Performance", "HUMANITIES", "bg-secondary text-white", "Detailed grade distribution for core and elective subjects in Humanities.")}
                    ${renderSubjSectionTable(sciList, "7. Science Stream - Subject Performance", "SCIENCE", "bg-primary text-white", "Detailed grade distribution for core and elective subjects in Science.")}
                    ${otherGroups.map((grpName, gIdx) => {
                        const grpSubjs = allSubjs.filter(it => it.group.toUpperCase() === grpName.toUpperCase() && it.type === "ALL");
                        return renderSubjSectionTable(grpSubjs, `${8 + gIdx}. ${grpName} Stream - Subject Performance`, grpName.toUpperCase(), "bg-secondary text-white");
                    }).join("")}
                `;
            } else {
                const filtered = allSubjs.filter(item => {
                    if (item.type !== targetType) return false;
                    if (selectedGroup !== "ALL" && item.group !== selectedGroup) return false;
                    if (query && !item.subject.toLowerCase().includes(query)) return false;
                    return true;
                });

                const totApp = filtered.reduce((a, b) => a + (b.appeared || 0), 0);
                const totPass = filtered.reduce((a, b) => a + (b.passed || 0), 0);
                const totFail = filtered.reduce((a, b) => a + (b.failed || 0), 0);
                const passRate = totApp > 0 ? ((totPass / totApp) * 100).toFixed(2) : "0.00";

                kpiHtml = `
                    <div class="kpi-card"><div class="kpi-label">Subjects Evaluated</div><div class="kpi-num text-dark">${filtered.length}</div></div>
                    <div class="kpi-card"><div class="kpi-label">Total Appeared</div><div class="kpi-num text-primary">${totApp}</div></div>
                    <div class="kpi-card kpi-success"><div class="kpi-label">Total Passed</div><div class="kpi-num text-success">${totPass}</div></div>
                    <div class="kpi-card kpi-danger"><div class="kpi-label">Total Failed</div><div class="kpi-num text-danger">${totFail}</div></div>
                    <div class="kpi-card kpi-warning"><div class="kpi-label">Overall Pass %</div><div class="kpi-num text-warning">${passRate}%</div></div>
                `;

                tableHtml = filtered.length > 0 ? renderSubjSectionTable(filtered, `${typeLabel} - ${selectedGroup === "ALL" ? "All Groups" : selectedGroup}`, targetType, "bg-dark text-white") : `<p class="text-muted fst-italic p-3 text-center">No subject records match the selected filter.</p>`;
            }

            generatePdfReportWindow(
                typeLabel,
                `Group: ${selectedGroup === "ALL" ? "All Groups Combined" : selectedGroup} &bull; Filter: ${query ? `"${query}"` : "Complete Evaluation"}`,
                meta,
                kpiHtml,
                tableHtml,
                "Official subject-level grade distribution and pass evaluation."
            );
            return;
        }

        if (optionKey === "merit") {
            const groupFilter = document.getElementById("meritGroupFilter");
            const searchInput = document.getElementById("meritSearch");
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const allStudents = (data.merit_list && data.merit_list.students) || [];
            const filtered = allStudents.filter(st => {
                if (selectedGroup !== "ALL" && st.Group !== selectedGroup) return false;
                if (query) {
                    const str = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            const goldenCount = filtered.filter(s => s.Is_GPA5_Without_4th).length;
            const gpa5Count = filtered.filter(s => s.Is_GPA5_With_4th).length;
            const topStudent = filtered.length > 0 ? filtered[0] : null;

            const kpiHtml = `
                <div class="kpi-card kpi-success"><div class="kpi-label">Ranked Passed Candidates</div><div class="kpi-num text-success">${filtered.length}</div></div>
                <div class="kpi-card kpi-warning"><div class="kpi-label">Golden A+ (Pure / No 4th)</div><div class="kpi-num text-warning">${goldenCount}</div></div>
                <div class="kpi-card"><div class="kpi-label">GPA 5.0 (With 4th Sub)</div><div class="kpi-num text-primary">${gpa5Count}</div></div>
                <div class="kpi-card kpi-dark"><div class="kpi-label">Merit Rank #1 Topper</div><div class="kpi-num text-dark fs-7">${topStudent ? `${topStudent.Name} (${topStudent.Roll})` : 'N/A'}</div></div>
            `;

            const tableHtml = filtered.length > 0 ? `
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th style="width: 55px;">Merit #</th>
                            <th style="width: 55px;">Group #</th>
                            <th style="width: 80px;">Roll No</th>
                            <th style="width: 90px;">Reg No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>GPA (With 4th)</th>
                            <th>GPA (Pure)</th>
                            <th>A+</th>
                            <th>A</th>
                            <th class="text-start ps-2">Passed Subjects Summary</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filtered.map(st => `
                            <tr>
                                <td class="fw-bold">${st.Merit_Position}</td>
                                <td>${st.Group_Merit_Position || st.Merit_Position}</td>
                                <td class="fw-bold">${st.Roll}</td>
                                <td class="font-monospace text-muted">${st.Registration || "-"}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${st.Name}</td>
                                <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                                <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                <td class="fw-bold ${st.Is_GPA5_With_4th ? 'badge-gpa-gold' : 'text-primary'}">${st.GPA_With_4th}</td>
                                <td class="${st.Is_GPA5_Without_4th ? 'badge-gpa-gold fw-bold' : 'text-muted'}">${st.GPA_Without_4th}</td>
                                <td class="badge-gpa-gold">${st.A_Plus_Count || 0}</td>
                                <td class="badge-gpa-green">${st.A_Count || 0}</td>
                                <td class="text-start ps-2 small text-muted text-truncate" style="max-width: 220px;">${st.Passed_Grades_Summary || "-"}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No merit students found matching criteria.</p>`;

            generatePdfReportWindow(
                "Student Academic Merit List",
                `Ranked in Order of Academic Excellence &bull; Group: ${selectedGroup === "ALL" ? "All Groups Combined" : selectedGroup}`,
                meta,
                kpiHtml,
                tableHtml,
                "Official student merit ranking based on GPA With 4th Subject, Pure GPA Without 4th Subject, A+ count, and academic tie-breakers."
            );
            return;
        }

        if (optionKey === "fail_subj_asc") {
            const groupFilter = document.getElementById("failSubjGroupFilter");
            const searchInput = document.getElementById("failSubjSearch");
            const selectedGroup = groupFilter ? groupFilter.value : "ALL";
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();

            const allEntries = (data.failed_list_subject_wise && data.failed_list_subject_wise.entries) || [];
            const filtered = allEntries.filter(item => {
                if (selectedGroup !== "ALL" && item.Group !== selectedGroup) return false;
                if (query) {
                    const str = `${item.Failed_Subject} ${item.Roll} ${item.Registration || ""} ${item.Name} ${item.Group} ${item.Type || ""}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            const uniqueFailedStudents = new Set(filtered.map(e => e.Roll)).size;

            const kpiHtml = `
                <div class="kpi-card kpi-danger"><div class="kpi-label">Subject Failure Entries</div><div class="kpi-num text-danger">${filtered.length}</div></div>
                <div class="kpi-card"><div class="kpi-label">Failed Candidates Count</div><div class="kpi-num text-dark">${uniqueFailedStudents}</div></div>
                <div class="kpi-card kpi-warning"><div class="kpi-label">Sort Order</div><div class="kpi-num text-warning fs-7">Subject Ascending (A-Z)</div></div>
            `;

            const tableHtml = filtered.length > 0 ? `
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th class="text-start ps-2">Failed Subject (F)</th>
                            <th style="width: 80px;">Roll No</th>
                            <th style="width: 90px;">Reg No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>Failed Count</th>
                            <th class="text-start ps-2">All Failed Subjects</th>
                            <th class="text-start ps-2">Passed Grades</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filtered.map(item => `
                            <tr>
                                <td class="text-start ps-2 fw-bold text-danger">${item.Failed_Subject}</td>
                                <td class="fw-bold">${item.Roll}</td>
                                <td class="font-monospace text-muted">${item.Registration || "-"}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${item.Name}</td>
                                <td><span class="badge bg-light text-secondary border">${item.Group}</span></td>
                                <td><span class="badge ${item.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${item.Type || "REGULAR"}</span></td>
                                <td class="fw-bold">${item.Failed_Count} Sub</td>
                                <td class="text-start ps-2 small text-danger">${Array.isArray(item.All_Failed_Subjects) ? item.All_Failed_Subjects.join(", ") : item.All_Failed_Subjects}</td>
                                <td class="text-start ps-2 small text-muted text-truncate" style="max-width: 180px;">${item.Passed_Grades_Summary || "-"}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No failed subject records found.</p>`;

            generatePdfReportWindow(
                "Subject-Wise Failed Students List (Ascending Order)",
                `Organized alphabetically by failed subject code & name &bull; Group: ${selectedGroup === "ALL" ? "All Groups Combined" : selectedGroup}`,
                meta,
                kpiHtml,
                tableHtml,
                "Lists all candidates who received an 'F' grade, organized by subject in ascending order for remedial review."
            );
            return;
        }

        if (optionKey === "1fail") {
            const oneFail = data.one_subject_fail || {};
            const searchInput = document.getElementById("oneSubSearch");
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();
            const students = (oneFail.students || []).filter(st => {
                if (query) {
                    const str = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${st.Failed_Subject}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            const kpiHtml = `
                <div class="kpi-card kpi-warning"><div class="kpi-label">1-Subject Fail Candidates</div><div class="kpi-num text-warning">${students.length}</div></div>
                <div class="kpi-card kpi-success"><div class="kpi-label">Supplementary Exam Eligible</div><div class="kpi-num text-success">${students.length} (100%)</div></div>
            `;

            const subjSummary = oneFail.subjects_summary || {};
            const subjEntries = Object.entries(subjSummary);
            const summaryTableHtml = subjEntries.length > 0 ? `
                <div style="margin-bottom: 16px;">
                    <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; color: #1e293b; letter-spacing: 0.04em;">
                        Summary: Number of Students Failed by Subject
                    </div>
                    <table class="table table-bordered table-sm text-center" style="margin-bottom: 15px;">
                        <thead>
                            <tr style="background-color: #f8fafc;">
                                <th class="text-start ps-2">Failed Subject</th>
                                <th style="width: 140px;">Failed Students Count</th>
                                <th style="width: 110px;">% Share</th>
                                <th style="width: 170px;">Supplementary Eligibility</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${subjEntries.map(([s, c]) => `
                                <tr>
                                    <td class="text-start ps-2 fw-bold text-danger">${s}</td>
                                    <td class="fw-bold">${c}</td>
                                    <td>${students.length > 0 ? ((c / students.length) * 100).toFixed(1) : 0}%</td>
                                    <td><span class="badge bg-success-subtle text-success border">Retake Eligible</span></td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>
            ` : "";

            const tableHtml = students.length > 0 ? `
                ${summaryTableHtml}
                <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; color: #1e293b; letter-spacing: 0.04em;">
                    Detailed Student Roster
                </div>
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th style="width: 80px;">Roll No</th>
                            <th style="width: 90px;">Reg No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>Failed Subject (F)</th>
                            <th class="text-start ps-2">Passed Subjects Summary</th>
                            <th>Eligibility Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${students.map(st => `
                            <tr>
                                <td class="fw-bold">${st.Roll}</td>
                                <td class="font-monospace text-muted">${st.Registration || "-"}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${st.Name}</td>
                                <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                                <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                <td class="fw-bold text-danger">${st.Failed_Subject}</td>
                                <td class="text-start ps-2 small text-muted">${st.Passed_Grades_Summary || "-"}</td>
                                <td><span class="badge bg-success-subtle text-success border border-success-subtle">Eligible for Supplementary</span></td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No students failed in only one subject.</p>`;

            generatePdfReportWindow(
                "One Subject Fail - Supplementary Eligible Candidates Roster",
                `Total: ${students.length} Students`,
                meta,
                kpiHtml,
                tableHtml,
                "These candidates failed in exactly ONE subject and are eligible to sit for remedial / supplementary examination to clear their diploma/certificate."
            );
            return;
        }

        if (optionKey === "2fail") {
            const twoFail = data.two_subject_fail || {};
            const searchInput = document.getElementById("twoSubSearch");
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();
            const students = (twoFail.students || []).filter(st => {
                if (query) {
                    const str = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""} ${(st.Failed_Subjects || []).join(" ")}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            const kpiHtml = `
                <div class="kpi-card kpi-danger"><div class="kpi-label">2-Subject Fail Candidates</div><div class="kpi-num text-danger">${students.length}</div></div>
            `;

            const subjs2Entries = Object.entries(twoFail.subjects_summary || {});
            const pairs2Entries = Object.entries(twoFail.pair_summary || {});

            const summary2Html = (subjs2Entries.length > 0 || pairs2Entries.length > 0) ? `
                <div style="display: flex; gap: 12px; margin-bottom: 16px;">
                    <div style="flex: 1;">
                        <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; color: #1e293b;">Top Failing Subjects (2-Sub Fails)</div>
                        <table class="table table-bordered table-sm text-center">
                            <thead><tr style="background-color: #f8fafc;"><th class="text-start ps-2">Subject</th><th style="width: 80px;">Count</th></tr></thead>
                            <tbody>${subjs2Entries.slice(0, 5).map(([s, c]) => `<tr><td class="text-start ps-2 fw-semibold text-danger">${s}</td><td>${c}</td></tr>`).join("")}</tbody>
                        </table>
                    </div>
                    <div style="flex: 1;">
                        <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; color: #1e293b;">Top Failing Subject Pairs (Combinations)</div>
                        <table class="table table-bordered table-sm text-center">
                            <thead><tr style="background-color: #f8fafc;"><th class="text-start ps-2">Subject Pair</th><th style="width: 80px;">Students</th></tr></thead>
                            <tbody>${pairs2Entries.slice(0, 5).map(([p, c]) => `<tr><td class="text-start ps-2 fw-semibold">${p}</td><td>${c}</td></tr>`).join("")}</tbody>
                        </table>
                    </div>
                </div>
            ` : "";

            const tableHtml = students.length > 0 ? `
                ${summary2Html}
                <div style="font-size: 11px; font-weight: bold; text-transform: uppercase; margin-bottom: 6px; color: #1e293b; letter-spacing: 0.04em;">
                    Detailed Student Roster
                </div>
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th style="width: 80px;">Roll No</th>
                            <th style="width: 90px;">Reg No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>Failed Subjects (2)</th>
                            <th class="text-start ps-2">Passed Subjects Summary</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${students.map(st => `
                            <tr>
                                <td class="fw-bold">${st.Roll}</td>
                                <td class="font-monospace text-muted">${st.Registration || "-"}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${st.Name}</td>
                                <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                                <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                <td class="fw-bold text-danger">${(st.Failed_Subjects || []).join(", ")}</td>
                                <td class="text-start ps-2 small text-muted">${st.Passed_Grades_Summary || "-"}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No students failed in two subjects.</p>`;

            generatePdfReportWindow(
                "Two Subject Fail Candidates List",
                `Total: ${students.length} Students`,
                meta,
                kpiHtml,
                tableHtml,
                "Candidates failing in exactly two subjects who require targeted academic tutoring in these specific subjects."
            );
            return;
        }

        if (optionKey === "mfail") {
            const mFail = data.multi_subject_fail || {};
            const students = mFail.students || [];

            const kpiHtml = `
                <div class="kpi-card kpi-dark"><div class="kpi-label">3+ Subject Fail Candidates</div><div class="kpi-num text-dark">${students.length}</div></div>
            `;

            const tableHtml = students.length > 0 ? `
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th style="width: 80px;">Roll No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>Failed Count</th>
                            <th>Failed Subjects</th>
                            <th class="text-start ps-2">Passed Subjects Summary</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${students.map(st => `
                            <tr>
                                <td class="fw-bold">${st.Roll}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${st.Name}</td>
                                <td><span class="badge bg-light text-secondary border">${st.Group}</span></td>
                                <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                <td class="fw-bold text-danger">${st.Failed_Count} Subjects</td>
                                <td class="text-danger fw-semibold">${(st.Failed_Subjects || []).join(", ")}</td>
                                <td class="text-start ps-2 small text-muted">${st.Passed_Grades_Summary || "-"}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No students failed in 3 or more subjects.</p>`;

            generatePdfReportWindow(
                "Multi-Subject Fail (3+ Subjects) Candidates Roster",
                `Total: ${students.length} Students`,
                meta,
                kpiHtml,
                tableHtml,
                "Candidates with multiple subject failures requiring targeted academic guidance."
            );
            return;
        }

        if (optionKey === "absent") {
            const absentData = data.absent_analysis || {};
            const searchInput = document.getElementById("absentSearch");
            const query = (searchInput ? searchInput.value : "").trim().toLowerCase();
            const students = (absentData.students || []).filter(st => {
                if (query) {
                    const str = `${st.Roll} ${st.Registration || ""} ${st.Name} ${st.Group} ${st.Type || ""}`.toLowerCase();
                    if (!str.includes(query)) return false;
                }
                return true;
            });

            const kpiHtml = `
                <div class="kpi-card kpi-danger"><div class="kpi-label">Total Absent Candidates</div><div class="kpi-num text-danger">${students.length}</div></div>
                <div class="kpi-card"><div class="kpi-label">Absent Rate %</div><div class="kpi-num text-dark">${absentData.absent_rate || 0}%</div></div>
                <div class="kpi-card"><div class="kpi-label">Regular Absent</div><div class="kpi-num text-primary">${absentData.regular_absent || 0}</div></div>
                <div class="kpi-card"><div class="kpi-label">Irregular Absent</div><div class="kpi-num text-warning">${absentData.irregular_absent || 0}</div></div>
                <div class="kpi-card kpi-success"><div class="kpi-label">Attended Examination</div><div class="kpi-num text-success">${absentData.appeared || 0}</div></div>
            `;

            const tableHtml = students.length > 0 ? `
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th style="width: 80px;">Roll No</th>
                            <th style="width: 90px;">Reg No</th>
                            <th class="text-start ps-2">Student Name</th>
                            <th class="text-start ps-2">Father's Name</th>
                            <th>Group</th>
                            <th>Type</th>
                            <th>Institute</th>
                            <th>Status</th>
                            <th>Remarks</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${students.map(st => `
                            <tr>
                                <td class="fw-bold">${st.Roll || "-"}</td>
                                <td class="font-monospace text-muted">${st.Registration || "-"}</td>
                                <td class="text-start ps-2 fw-semibold text-dark">${st.Name || "N/A"}</td>
                                <td class="text-start ps-2 text-muted small">${st["Father Name"] || st.Father_Name || "-"}</td>
                                <td><span class="badge bg-light text-secondary border">${st.Group || "GENERAL"}</span></td>
                                <td><span class="badge ${st.Type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${st.Type || "REGULAR"}</span></td>
                                <td class="small text-truncate" style="max-width: 150px;">${st.Institute || "-"}</td>
                                <td><span class="badge bg-secondary">ABSENT</span></td>
                                <td class="small text-muted">${st.Remarks || "Did not sit for examination"}</td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            ` : `<p class="text-muted fst-italic p-3 text-center">No absent candidates recorded.</p>`;

            generatePdfReportWindow(
                "Absent Candidates Examination Roster",
                `Total Absent: ${students.length} Candidates (${absentData.absent_rate || 0}% of Enrolled)`,
                meta,
                kpiHtml,
                tableHtml,
                "Official board attendance verification roster of candidates who did not sit for the examination."
            );
            return;
        }

        if (optionKey === "summary") {
            const groupsSummary = data.groups_summary || {};
            const gpaDist = data.gpa_distribution_with_4th || data.gpa_distribution || {};
            const summary = data.summary || {};

            const kpiHtml = `
                <div class="kpi-card"><div class="kpi-label">Total Appeared</div><div class="kpi-num text-primary">${summary.appeared || 0}</div></div>
                <div class="kpi-card kpi-success"><div class="kpi-label">Total Passed</div><div class="kpi-num text-success">${summary.passed || 0}</div></div>
                <div class="kpi-card kpi-danger"><div class="kpi-label">Total Failed</div><div class="kpi-num text-danger">${summary.failed || 0}</div></div>
                <div class="kpi-card kpi-warning"><div class="kpi-label">Total GPA 5.0</div><div class="kpi-num text-warning">${summary.gpa_5_count_with_4th || summary.gpa_5_count || 0}</div></div>
            `;

            const groupRows = Object.values(groupsSummary).map(grp => `
                <tr>
                    <td class="text-start ps-2 fw-bold text-dark">${grp.group}</td>
                    <td>${grp.total}</td>
                    <td>${grp.appeared}</td>
                    <td class="text-muted">${grp.absent || 0}</td>
                    <td class="text-success fw-bold">${grp.passed}</td>
                    <td class="text-danger fw-bold">${grp.failed}</td>
                    <td class="fw-bold">${grp.pass_rate}%</td>
                    <td class="badge-gpa-gold">${grp.gpa5_count_with_4th !== undefined ? grp.gpa5_count_with_4th : grp.gpa5_count}</td>
                    <td class="text-primary fw-bold">${grp.gpa5_count_without_4th !== undefined ? grp.gpa5_count_without_4th : 0}</td>
                    <td>${grp.one_fail_count}</td>
                    <td>${grp.two_fail_count}</td>
                    <td>${grp.multi_fail_count}</td>
                </tr>
            `).join("");

            const gpaRows = Object.entries(gpaDist).map(([tier, cnt]) => `
                <tr>
                    <td class="fw-bold text-start ps-2">${tier}</td>
                    <td class="fw-bold">${cnt}</td>
                    <td>${summary.appeared > 0 ? ((cnt / summary.appeared) * 100).toFixed(2) : '0.00'}%</td>
                </tr>
            `).join("");

            const contentHtml = `
                <h6 class="fw-bold text-dark text-uppercase mt-2 mb-2">Group Performance Breakdown</h6>
                <table class="table table-bordered table-sm mb-4">
                    <thead>
                        <tr>
                            <th class="text-start ps-2">Group</th>
                            <th>Total</th>
                            <th>Appeared</th>
                            <th>Absent</th>
                            <th>Passed</th>
                            <th>Failed</th>
                            <th>Pass %</th>
                            <th class="badge-gpa-gold">GPA 5 (With 4th)</th>
                            <th>GPA 5 (Without 4th)</th>
                            <th>1-Fail</th>
                            <th>2-Fail</th>
                            <th>3+ Fail</th>
                        </tr>
                    </thead>
                    <tbody>${groupRows}</tbody>
                </table>

                <h6 class="fw-bold text-dark text-uppercase mt-4 mb-2">GPA Tier Distribution</h6>
                <table class="table table-bordered table-sm" style="max-width: 500px;">
                    <thead>
                        <tr>
                            <th class="text-start ps-2">GPA Grade Tier</th>
                            <th>Student Count</th>
                            <th>Percentage of Appeared</th>
                        </tr>
                    </thead>
                    <tbody>${gpaRows}</tbody>
                </table>
            `;

            generatePdfReportWindow(
                "Group Performance Comparison & GPA Breakdown Report",
                `Analysis by academic curriculum group and GPA tier distribution`,
                meta,
                kpiHtml,
                contentHtml
            );
            return;
        }

        if (optionKey === "reg_irreg") {
            const reg = data.regular_summary || {};
            const irreg = data.irregular_summary || {};
            const typesSummary = data.types_summary || {};
            const groupType = data.group_type_comparison || [];

            const kpiHtml = `
                <div class="kpi-card"><div class="kpi-label">Regular Students</div><div class="kpi-num text-primary">${reg.total || 0}</div><div class="small fw-bold text-success">${reg.pass_rate || 0}% Pass</div></div>
                <div class="kpi-card kpi-warning"><div class="kpi-label">Irregular Students</div><div class="kpi-num text-warning">${irreg.total || 0}</div><div class="small fw-bold text-dark">${irreg.pass_rate || 0}% Pass</div></div>
                <div class="kpi-card kpi-success"><div class="kpi-label">Regular GPA 5.0</div><div class="kpi-num text-success">${reg.gpa5_count || 0}</div></div>
                <div class="kpi-card kpi-danger"><div class="kpi-label">Irregular Failures</div><div class="kpi-num text-danger">${irreg.failed || 0}</div></div>
            `;

            const typeRows = Object.values(typesSummary).map(t => `
                <tr>
                    <td class="text-start ps-2 fw-bold">${t.type}</td>
                    <td>${t.total}</td>
                    <td>${t.appeared}</td>
                    <td class="text-success fw-bold">${t.passed}</td>
                    <td class="text-danger fw-bold">${t.failed}</td>
                    <td class="fw-bold">${t.pass_rate}%</td>
                    <td class="badge-gpa-gold">${t.gpa5_count}</td>
                    <td>${t.one_fail_count}</td>
                    <td>${t.two_fail_count}</td>
                    <td>${t.multi_fail_count}</td>
                </tr>
            `).join("");

            const matrixRows = groupType.map(row => `
                <tr>
                    <td class="text-start ps-2 fw-semibold">${row.group}</td>
                    <td><span class="badge ${row.type === 'REGULAR' ? 'bg-primary-subtle text-primary border' : 'bg-warning-subtle text-dark border'}">${row.type}</span></td>
                    <td>${row.total}</td>
                    <td>${row.appeared}</td>
                    <td class="text-success fw-bold">${row.passed}</td>
                    <td class="text-danger fw-bold">${row.failed}</td>
                    <td class="fw-bold">${row.pass_rate}%</td>
                    <td class="badge-gpa-gold">${row.gpa5_count}</td>
                    <td>${row.one_fail_count}</td>
                    <td>${row.two_fail_count}</td>
                    <td>${row.multi_fail_count}</td>
                </tr>
            `).join("");

            const contentHtml = `
                <h6 class="fw-bold text-dark text-uppercase mt-2 mb-2">Overall Candidate Type Performance</h6>
                <table class="table table-bordered table-sm mb-4">
                    <thead>
                        <tr>
                            <th class="text-start ps-2">Candidate Category</th>
                            <th>Total</th>
                            <th>Appeared</th>
                            <th>Passed</th>
                            <th>Failed</th>
                            <th>Pass Rate</th>
                            <th class="badge-gpa-gold">GPA 5.0</th>
                            <th>1-Sub Fail</th>
                            <th>2-Sub Fail</th>
                            <th>3+ Sub Fail</th>
                        </tr>
                    </thead>
                    <tbody>${typeRows}</tbody>
                </table>

                <h6 class="fw-bold text-dark text-uppercase mt-4 mb-2">Group x Candidate Type Matrix</h6>
                <table class="table table-bordered table-sm">
                    <thead>
                        <tr>
                            <th class="text-start ps-2">Group</th>
                            <th>Candidate Category</th>
                            <th>Total</th>
                            <th>Appeared</th>
                            <th>Passed</th>
                            <th>Failed</th>
                            <th>Pass Rate</th>
                            <th class="badge-gpa-gold">GPA 5.0</th>
                            <th>1-Sub Fail</th>
                            <th>2-Sub Fail</th>
                            <th>3+ Sub Fail</th>
                        </tr>
                    </thead>
                    <tbody>${matrixRows}</tbody>
                </table>
            `;

            generatePdfReportWindow(
                "Regular vs Irregular Candidates Comparative Analysis",
                `Comparison of regular vs irregular candidates performance`,
                meta,
                kpiHtml,
                contentHtml
            );
            return;
        }
    }
});
