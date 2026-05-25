<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SQL Generator - VTTU</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://unpkg.com/lucide@latest"></script>
</head>
<body class="bg-slate-50 min-h-screen">
    <div class="container mx-auto py-10 px-4">
        <div class="max-w-4xl mx-auto">
            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <!-- Header -->
                <div class="border-b border-slate-200 bg-white px-6 py-4">
                    <h1 class="text-xl font-semibold text-slate-800 flex items-center gap-2">
                        <i data-lucide="database" class="w-5 h-5 text-blue-600"></i>
                        VTTU - SQL Generator
                    </h1>
                </div>

                <!-- Tabs -->
                <div class="flex border-b border-slate-200 bg-slate-50/50">
                    <button onclick="switchTab('dv')" id="tab-dv" class="px-6 py-3 text-sm font-medium border-b-2 border-blue-600 text-blue-600 bg-white transition-all">
                        AddMaThuocDV
                    </button>
                    <button onclick="switchTab('kigui')" id="tab-kigui" class="px-6 py-3 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-all">
                        AddMaThuocKiGui
                    </button>
                    <button onclick="switchTab('vtytdv')" id="tab-vtytdv" class="px-6 py-3 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-all">
                        AddMaVTYTDV
                    </button>
                    <button onclick="switchTab('vtytkigui')" id="tab-vtytkigui" class="px-6 py-3 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 transition-all">
                        AddMaVTYTKIGUI
                    </button>
                </div>

                <!-- Content -->
                <div class="p-6">
                    <form id="uploadForm" class="space-y-6">
                        @csrf
                        <input type="hidden" name="type" id="processType" value="dv">
                        <div class="grid grid-cols-1 gap-6">
                            <div class="border-2 border-dashed border-slate-300 rounded-lg p-12 text-center hover:border-blue-400 transition-colors cursor-pointer" id="dropZone">
                                <input type="file" name="excel_file" id="excelFile" class="hidden" accept=".xlsx,.xls,.csv">
                                <div class="flex flex-col items-center gap-2">
                                    <i data-lucide="upload-cloud" class="w-10 h-10 text-slate-400"></i>
                                    <p class="text-slate-600 font-medium" id="fileNameDisplay">Kéo thả file Excel vào đây hoặc click để chọn</p>
                                    <p class="text-xs text-slate-400 uppercase">Hỗ trợ: .xlsx, .xls, .csv</p>
                                </div>
                            </div>

                            <div class="flex justify-end">
                                <button type="submit" id="submitBtn" class="bg-blue-600 hover:bg-blue-700 text-white px-6 py-2 rounded-lg font-medium transition-all flex items-center gap-2 disabled:opacity-50">
                                    <i data-lucide="play" class="w-4 h-4"></i>
                                    Tạo câu lệnh SQL
                                </button>
                            </div>
                        </div>
                    </form>

                    <!-- Result Section -->
                    <div id="resultSection" class="mt-8 hidden space-y-8">
                        <!-- Part 1: Main Excel Data -->
                        <div class="space-y-3">
                            <div class="flex items-center justify-between">
                                <h3 class="text-sm font-semibold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                                    <span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">1</span>
                                    Bước 1: Insert dữ liệu từ Excel
                                </h3>
                                <button onclick="copyToClipboard('sqlOutput', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md border border-slate-300 transition-all flex items-center gap-1">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                    Sao chép
                                </button>
                            </div>
                            <div class="relative">
                                <textarea id="sqlOutput" readonly class="w-full h-80 p-4 bg-slate-900 text-emerald-400 font-mono text-sm rounded-lg border border-slate-800 focus:outline-none"></textarea>
                            </div>
                        </div>

                        <!-- Part 2: Init -->
                        <div class="space-y-3">
                            <div class="flex items-center justify-between">
                                <h3 class="text-sm font-semibold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                                    <span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">2</span>
                                    Bước 2: Khởi tạo bảng tạm & Tồn kho
                                </h3>
                                <button onclick="copyToClipboard('sqlInit', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md border border-slate-300 transition-all flex items-center gap-1">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                    Sao chép
                                </button>
                            </div>
                            <div class="relative">
                                <textarea id="sqlInit" readonly class="w-full h-32 p-4 bg-slate-900 text-blue-400 font-mono text-sm rounded-lg border border-slate-800 focus:outline-none"></textarea>
                            </div>
                        </div>

                        <!-- Part 3: Warning/Config -->
                        <div class="space-y-3">
                            <div class="flex items-center justify-between">
                                <h3 class="text-sm font-semibold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                                    <span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">3</span>
                                    Bước 3: Cập nhật Cảnh báo tồn
                                </h3>
                                <button onclick="copyToClipboard('sqlWarning', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md border border-slate-300 transition-all flex items-center gap-1">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                    Sao chép
                                </button>
                            </div>
                            <div class="relative">
                                <textarea id="sqlWarning" readonly class="w-full h-32 p-4 bg-slate-900 text-orange-400 font-mono text-sm rounded-lg border border-slate-800 focus:outline-none"></textarea>
                            </div>
                        </div>

                        <!-- Part 4: Kigui Update (Only for Kigui) -->
                        <div id="step4Container" class="space-y-3 hidden">
                            <div class="flex items-center justify-between">
                                <h3 class="text-sm font-semibold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                                    <span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">4</span>
                                    Bước 4: Cập nhật Nhà thuốc ngoài (Kí gửi)
                                </h3>
                                <button onclick="copyToClipboard('sqlKiguiUpdate', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md border border-slate-300 transition-all flex items-center gap-1">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                    Sao chép
                                </button>
                            </div>
                            <div class="relative">
                                <textarea id="sqlKiguiUpdate" readonly class="w-full h-32 p-4 bg-slate-900 text-purple-400 font-mono text-sm rounded-lg border border-slate-800 focus:outline-none"></textarea>
                            </div>
                        </div>

                        <!-- Part 5: Kigui Check (Only for Kigui) -->
                        <div id="step5Container" class="space-y-3 hidden">
                            <div class="flex items-center justify-between">
                                <h3 class="text-sm font-semibold text-slate-700 uppercase tracking-wider flex items-center gap-2">
                                    <span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">5</span>
                                    Bước 5: Kiểm tra dữ liệu (Optional)
                                </h3>
                                <button onclick="copyToClipboard('sqlCheck', this)" class="text-xs bg-slate-100 hover:bg-slate-200 text-slate-700 px-3 py-1.5 rounded-md border border-slate-300 transition-all flex items-center gap-1">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                    Sao chép
                                </button>
                            </div>
                            <div class="relative">
                                <textarea id="sqlCheck" readonly class="w-full h-32 p-4 bg-slate-900 text-pink-400 font-mono text-sm rounded-lg border border-slate-800 focus:outline-none"></textarea>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Note -->
            <div class="mt-6 p-4 bg-blue-50 border border-blue-100 rounded-lg">
                <div class="flex gap-3">
                    <i data-lucide="info" class="w-5 h-5 text-blue-500 shrink-0"></i>
                    <div class="text-sm text-blue-700">
                        <p class="font-semibold mb-1">Hướng dẫn:</p>
                        <ul class="list-disc list-inside space-y-1 opacity-90">
                            <li>Chọn đúng Tab (Dịch vụ hoặc Kí gửi) trước khi upload file.</li>
                            <li>File Excel dữ liệu bắt đầu từ dòng 3.</li>
                            <li>Kết quả SQL sẽ được chia thành 3 bước để bạn thực thi tuần tự.</li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        lucide.createIcons();

        let currentTab = 'dv';
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('excelFile');
        const fileNameDisplay = document.getElementById('fileNameDisplay');
        const uploadForm = document.getElementById('uploadForm');
        const resultSection = document.getElementById('resultSection');
        const sqlInit = document.getElementById('sqlInit');
        const sqlOutput = document.getElementById('sqlOutput');
        const sqlWarning = document.getElementById('sqlWarning');
        const sqlKiguiUpdate = document.getElementById('sqlKiguiUpdate');
        const sqlCheck = document.getElementById('sqlCheck');
        const submitBtn = document.getElementById('submitBtn');
        const processTypeInput = document.getElementById('processType');
        const step4Container = document.getElementById('step4Container');
        const step5Container = document.getElementById('step5Container');

        function switchTab(tab) {
            currentTab = tab;
            processTypeInput.value = tab;
            
            // UI Update
            const tabDv = document.getElementById('tab-dv');
            const tabKigui = document.getElementById('tab-kigui');
            const tabVtyt = document.getElementById('tab-vtytdv');
            const tabVtytkigui = document.getElementById('tab-vtytkigui');
            const step1Title = document.querySelector('#resultSection > div:nth-child(1) h3');
            const step2Title = document.querySelector('#resultSection > div:nth-child(2) h3');
            const step3Title = document.querySelector('#resultSection > div:nth-child(3) h3');
            const step4Title = document.querySelector('#step4Container h3');
            
            [tabDv, tabKigui, tabVtyt, tabVtytkigui].forEach(t => {
                t.classList.add('border-transparent', 'text-slate-500');
                t.classList.remove('border-blue-600', 'text-blue-600', 'bg-white');
            });

            const activeTab = document.getElementById('tab-' + tab);
            activeTab.classList.remove('border-transparent', 'text-slate-500');
            activeTab.classList.add('border-blue-600', 'text-blue-600', 'bg-white');

            if (tab === 'vtytdv' || tab === 'vtytkigui') {
                step1Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">1</span> Bước 1: Insert dữ liệu VTYT từ Excel';
                step2Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">2</span> Bước 2: Khởi tạo bảng tạm VTYT & Tồn kho';
                step3Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">3</span> Bước 3: Cập nhật Cảnh báo tồn VTYT';
                if (tab === 'vtytkigui') {
                    step4Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">4</span> Bước 4: Cập nhật Nhà thuốc ngoài (VTYT Kí gửi)';
                }
            } else {
                step1Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">1</span> Bước 1: Insert dữ liệu từ Excel';
                step2Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">2</span> Bước 2: Khởi tạo bảng tạm & Tồn kho';
                step3Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">3</span> Bước 3: Cập nhật Cảnh báo tồn';
                if (tab === 'kigui') {
                    step4Title.innerHTML = '<span class="bg-blue-600 text-white w-5 h-5 rounded-full flex items-center justify-center text-[10px]">4</span> Bước 4: Cập nhật Nhà thuốc ngoài (Kí gửi)';
                }
            }

            if (tab === 'dv' || tab === 'vtytdv') {
                step4Container.classList.add('hidden');
                step5Container.classList.add('hidden');
            }
            
            // Reset results
            resultSection.classList.add('hidden');
            fileNameDisplay.textContent = 'Kéo thả file Excel vào đây hoặc click để chọn';
            fileNameDisplay.classList.remove('text-blue-600');
            fileInput.value = '';
        }

        dropZone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                fileNameDisplay.textContent = e.target.files[0].name;
                fileNameDisplay.classList.add('text-blue-600');
            }
        });

        uploadForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const formData = new FormData(uploadForm);
            
            if (!fileInput.files[0]) {
                alert('Vui lòng chọn file!');
                return;
            }

            submitBtn.disabled = true;
            submitBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> Đang xử lý...';
            lucide.createIcons();

            let endpoint = '/add-ma-thuoc-dv/process';
            if (currentTab === 'kigui') endpoint = '/add-ma-thuoc-kigui/process';
            if (currentTab === 'vtytdv') endpoint = '/add-ma-vtyt-dv/process';
            if (currentTab === 'vtytkigui') endpoint = '/add-ma-vtyt-kigui/process';

            try {
                const response = await fetch(endpoint, {
                    method: 'POST',
                    body: formData,
                    headers: {
                        'X-CSRF-TOKEN': document.querySelector('input[name="_token"]').value
                    }
                });
                
                const result = await response.json();
                
                if (result.success) {
                    sqlInit.value = result.init_queries;
                    sqlOutput.value = result.queries;
                    sqlWarning.value = result.warning_queries;
                    
                    if (currentTab === 'kigui' || currentTab === 'vtytkigui') {
                        sqlKiguiUpdate.value = result.update_kigui_queries;
                        sqlCheck.value = result.check_queries;
                        step4Container.classList.remove('hidden');
                        step5Container.classList.remove('hidden');
                    } else {
                        step4Container.classList.add('hidden');
                        step5Container.classList.add('hidden');
                    }
                    
                    resultSection.classList.remove('hidden');
                    window.scrollTo({ top: resultSection.offsetTop - 20, behavior: 'smooth' });
                } else {
                    alert('Có lỗi xảy ra: ' + (result.message || 'Không rõ nguyên nhân'));
                }
            } catch (error) {
                console.error(error);
                alert('Lỗi kết nối server!');
            } finally {
                submitBtn.disabled = false;
                submitBtn.innerHTML = '<i data-lucide="play" class="w-4 h-4"></i> Tạo câu lệnh SQL';
                lucide.createIcons();
            }
        });

        async function copyToClipboard(elementId, btn) {
            const textArea = document.getElementById(elementId);
            try {
                await navigator.clipboard.writeText(textArea.value);
                const originalText = btn.innerHTML;
                btn.innerHTML = '<i data-lucide="check" class="w-3.5 h-3.5"></i> Đã chép!';
                btn.classList.replace('bg-slate-100', 'bg-green-100');
                btn.classList.replace('text-slate-700', 'text-green-700');
                lucide.createIcons();
                
                setTimeout(() => {
                    btn.innerHTML = originalText;
                    btn.classList.replace('bg-green-100', 'bg-slate-100');
                    btn.classList.replace('text-green-700', 'text-slate-700');
                    lucide.createIcons();
                }, 2000);
            } catch (err) {
                alert('Không thể copy!');
            }
        }
    </script>
</body>
</html>
