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

    // ==========================================
    // KNOWN INSTITUTIONS & BOARD DEFAULTS REGISTRY
    // ==========================================
    const KNOWN_INSTITUTIONS = {
        "132092": {
            name: "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE",
            address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "132092",
            board: "Cumilla"
        },
        "105820": {
            name: "COMILLA AJIT GUHA COLLEGE",
            address: "Chawkbazar, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "105820",
            board: "Cumilla"
        },
        "105814": {
            name: "CUMILLA VICTORIA GOVT. COLLEGE",
            address: "Victoria College Road, Cumilla Sadar, Cumilla - 3500, Bangladesh",
            eiin: "105814",
            board: "Cumilla"
        },
        "105786": {
            name: "CUMILLA ZILLA SCHOOL",
            address: "Kandirpar, Cumilla Sadar, Cumilla, Bangladesh",
            eiin: "105786",
            board: "Cumilla"
        },
        "107977": {
            name: "DHAKA COLLEGE",
            address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh",
            eiin: "107977",
            board: "Dhaka"
        },
        "107975": {
            name: "DHAKA CITY COLLEGE",
            address: "Road No 2, Dhanmondi, Dhaka - 1205, Bangladesh",
            eiin: "107975",
            board: "Dhaka"
        },
        "108274": {
            name: "NOTRE DAME COLLEGE",
            address: "Toyenbee Circular Road, Motijheel, Dhaka - 1000, Bangladesh",
            eiin: "108274",
            board: "Dhaka"
        },
        "108573": {
            name: "RAJUK UTTARA MODEL COLLEGE",
            address: "Sector-6, Uttara Model Town, Dhaka - 1230, Bangladesh",
            eiin: "108573",
            board: "Dhaka"
        },
        "107845": {
            name: "ADAMJEE CANTONMENT COLLEGE",
            address: "Dhaka Cantonment, Dhaka - 1206, Bangladesh",
            eiin: "107845",
            board: "Dhaka"
        },
        "108348": {
            name: "VIQARUNNISA NOON COLLEGE",
            address: "1/A Bailey Road, Ramna, Dhaka - 1000, Bangladesh",
            eiin: "108348",
            board: "Dhaka"
        },
        "108272": {
            name: "HOLY CROSS COLLEGE",
            address: "Tejgaon, Dhaka - 1215, Bangladesh",
            eiin: "108272",
            board: "Dhaka"
        },
        "108035": {
            name: "GOVT. SCIENCE COLLEGE",
            address: "Farmgate, Tejgaon, Dhaka - 1215, Bangladesh",
            eiin: "108035",
            board: "Dhaka"
        },
        "132140": {
            name: "MOHAMMADPUR GOVT. COLLEGE",
            address: "Satmasjid Road, Mohammadpur, Dhaka - 1207, Bangladesh",
            eiin: "132140",
            board: "Dhaka"
        },
        "133008": {
            name: "CHITTAGONG COLLEGE",
            address: "College Road, Chawkbazar, Chattogram - 4203, Bangladesh",
            eiin: "133008",
            board: "Chattogram"
        },
        "126490": {
            name: "RAJSHAHI COLLEGE",
            address: "Boalia, Rajshahi - 6000, Bangladesh",
            eiin: "126490",
            board: "Rajshahi"
        },
        "100742": {
            name: "GOVT. BRAJOMOHUN (BM) COLLEGE",
            address: "College Road, Barishal - 8200, Bangladesh",
            eiin: "100742",
            board: "Barishal"
        },
        "130456": {
            name: "MURARI CHAND (MC) COLLEGE",
            address: "Tilagarh, Sylhet - 3100, Bangladesh",
            eiin: "130456",
            board: "Sylhet"
        },
        "127485": {
            name: "CARMICHAEL COLLEGE",
            address: "Lalbag, Rangpur - 5400, Bangladesh",
            eiin: "127485",
            board: "Rangpur"
        },
        "117106": {
            name: "GOVT. BRAJALAL (BL) COLLEGE",
            address: "Daulatpur, Khulna - 9202, Bangladesh",
            eiin: "117106",
            board: "Khulna"
        },
        "111831": {
            name: "GOVT. ANANDA MOHAN COLLEGE",
            address: "College Road, Mymensingh - 2200, Bangladesh",
            eiin: "111831",
            board: "Mymensingh"
        },
        "106579": {
            name: "FENI GOVT. COLLEGE",
            address: "College Road, Feni - 3900, Bangladesh",
            eiin: "106579",
            board: "Cumilla"
        },
        "103284": {
            name: "BRAHMANBARIA GOVT. COLLEGE",
            address: "College Road, Brahmanbaria - 3400, Bangladesh",
            eiin: "103284",
            board: "Cumilla"
        },
        "103597": {
            name: "CHANDPUR GOVT. COLLEGE",
            address: "Stadium Road, Chandpur - 3600, Bangladesh",
            eiin: "103597",
            board: "Cumilla"
        },
        "107541": {
            name: "NOAKHALI GOVT. COLLEGE",
            address: "Maijdee Court, Noakhali - 3800, Bangladesh",
            eiin: "107541",
            board: "Cumilla"
        },
        "107386": {
            name: "BANDERHAT HIGH SCHOOL & COLLEGE",
            address: "Banderhat, Lakshmipur, Bangladesh",
            eiin: "107386",
            board: "Cumilla"
        },
        "115957": {
            name: "GOVT. M. M. COLLEGE, JASHORE",
            address: "Kharki, Jashore - 7400, Bangladesh",
            eiin: "115957",
            board: "Jessore"
        },
        "120583": {
            name: "DINAJPUR GOVT. COLLEGE",
            address: "Sadar, Dinajpur - 5200, Bangladesh",
            eiin: "120583",
            board: "Dinajpur"
        },
        "108156": {
            name: "GOVT. MADRASA-E-ALIA DHAKA",
            address: "Bakshibazar, Dhaka - 1211, Bangladesh",
            eiin: "108156",
            board: "Madrasah"
        },
        "108036": {
            name: "DHAKA POLYTECHNIC INSTITUTE",
            address: "Tejgaon I/A, Dhaka - 1208, Bangladesh",
            eiin: "108036",
            board: "Technical"
        }
    };

    const BOARD_DEFAULTS = {
        "dhaka": { name: "DHAKA COLLEGE", address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh", eiin: "107977" },
        "cumilla": { name: "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE", address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "132092" },
        "comilla": { name: "CUMILLA SHIKKHA BOARD GOVT. MODEL COLLEGE", address: "Kotbari Road, Cumilla Sadar, Cumilla - 3500, Bangladesh", eiin: "132092" },
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
        "technical": { name: "DHAKA POLYTECHNIC INSTITUTE", address: "Tejgaon I/A, Dhaka - 1208, Bangladesh", eiin: "108036" },
        "dibs": { name: "DHAKA COLLEGE (DIBS)", address: "Mirpur Road, Dhanmondi, Dhaka - 1205, Bangladesh", eiin: "107977" }
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

    function resolveInstitutionInfo(rawInstitute, rawEiin, boardName, rawAddress) {
        const b = (boardName || (boardSelect ? boardSelect.value : "dhaka")).toLowerCase().trim();
        const bClean = b.replace(/\(.*\)/, "").trim();
        const defaultForBoard = BOARD_DEFAULTS[bClean] || BOARD_DEFAULTS["dhaka"];

        let eiin = (rawEiin || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(eiin.toLowerCase())) eiin = "";
        let instName = (rawInstitute || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(instName.toLowerCase())) instName = "";
        let address = (rawAddress || "").toString().trim();
        if (["nan", "none", "n/a", "-", "null", "undefined"].includes(address.toLowerCase())) address = "";

        // 1. Check if rawInstitute has embedded digits like "(105963) ..." or "105963 - ..."
        if (instName) {
            const m = instName.match(/\(?\b(\d{5,7})\b\)?/);
            if (m) {
                if (!eiin) eiin = m[1];
                instName = instName.replace(/\(?\b\d{5,7}\b\)?/g, "").replace(/[-()]/g, " ").replace(/\s+/g, " ").trim();
            }
        }

        // 1b. Check per-institution saved EIIN override if still empty
        if (!eiin && instName) {
            eiin = getSavedEiinOverride(instName);
        }

        // 2. If EIIN is known in local curated map, get official address (never overwrite explicit different institution name)
        if (eiin && KNOWN_INSTITUTIONS[eiin]) {
            const match = KNOWN_INSTITUTIONS[eiin];
            return {
                name: instName || match.name,
                eiin: eiin,
                address: address || match.address
            };
        }

        // 3. Match exact curated institution name if no EIIN yet
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

        // 4. District matching for address
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
        if (!address && instName) {
            for (const dist of bdDistricts) {
                if (new RegExp(`\\b${dist}\\b`, "i").test(instName)) {
                    address = `${dist}, Bangladesh`;
                    break;
                }
            }
        }

        // 5. Only fall back to board default when NO institution name was provided at all (e.g. initial board preview)
        if (!instName) {
            if (!eiin) eiin = defaultForBoard.eiin;
            instName = defaultForBoard.name;
            if (!address) address = defaultForBoard.address;
        } else if (!address) {
            const titleBoard = bClean ? (bClean.charAt(0).toUpperCase() + bClean.slice(1)) : "Bangladesh";
            address = titleBoard === "Bangladesh" ? "Bangladesh" : `${titleBoard}, Bangladesh`;
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
        if (marksheetHeaderAddress) marksheetHeaderAddress.textContent = info.address || "Bangladesh";
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

    // Dynamic listener on EIIN editing: looks up full Bangladesh EIIN database via API
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
        const instName = document.getElementById("marksheetHeaderInstName")?.textContent?.trim() || "Educational Institution";
        const rawEiin = document.getElementById("marksheetHeaderEiin")?.textContent?.trim() || "";
        const eiin = (rawEiin && rawEiin.toUpperCase() !== "N/A" && rawEiin.toUpperCase() !== "EIIN") ? rawEiin : "";
        const studentName = document.getElementById("marksheetName")?.textContent?.trim() || "Student";
        const roll = document.getElementById("marksheetRoll")?.textContent?.trim() || "";
        const exam = document.getElementById("marksheetExam")?.textContent?.trim() || "Marksheet";
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
            const selectedBoard = boardSelect.value || "dhaka";
            if (activeBoardName) {
                const text = boardSelect.options[boardSelect.selectedIndex]?.text || "Dhaka";
                activeBoardName.textContent = text.replace(/\(.*\)/, "").trim();
            }
            // Dynamically update institutional header for selected board
            const info = resolveInstitutionInfo("", "", selectedBoard);
            updateMarksheetInstitutionalHeader(info);
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
        // Dynamically resolve Institutional Header Details (Name, Address, EIIN)
        const info = resolveInstitutionInfo(
            student.Institute,
            student.EIIN || student.eiin,
            student.Board || (boardSelect ? boardSelect.value : "Dhaka"),
            student.Institution_Address || student.address || ""
        );
        updateMarksheetInstitutionalHeader(info);
        if (!info.eiin && student.Institute) {
            fetch(`/api/lookup-institution?institute=${encodeURIComponent(student.Institute)}&board=${encodeURIComponent(student.Board || "")}`)
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
        marksheetBoard.textContent = (student.Board || (boardSelect ? boardSelect.value : "Dhaka")).toUpperCase();
        marksheetRoll.textContent = student.Roll || "-";
        marksheetReg.textContent = student.Registration || "-";
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

                // Set Download Links & Safe Filenames
                downloadXlsxBtn.href = msg.download_xlsx;
                downloadCsvBtn.href = msg.download_csv;
                if (msg.excel_filename) downloadXlsxBtn.setAttribute("download", msg.excel_filename);
                if (msg.csv_filename) downloadCsvBtn.setAttribute("download", msg.csv_filename);
                downloadSection.style.setProperty("display", "flex", "important");

                // Store analysis data for Result Analysis page
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
