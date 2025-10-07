document.addEventListener("DOMContentLoaded", async () => {
    // --- UTILITY FUNCTIONS ---
    const $ = id => document.getElementById(id);

    const showNotification = (message, type) => {
        const notification = $("notification");
        notification.textContent = message;
        notification.className = type;
        notification.style.display = 'block';
        setTimeout(() => { notification.style.display = 'none'; }, 5000);
    };

    /**
     * Calculates and assigns ranks to a list of items based on a score field.
     * Handles ties correctly (e.g., 99, 98, 98, 95 -> ranks 1, 2, 2, 4).
     * @param {Array<Object>} items - The array of objects to rank.
     * @param {string} scoreField - The property name of the score.
     * @param {string} rankField - The property name to store the calculated rank.
     */
    const assignRanks = (items, scoreField, rankField) => {
        const sortedItems = [...items].sort((a, b) => (b[scoreField] || -1) - (a[scoreField] || -1));
        if (sortedItems.length === 0) return;

        let rank = 1;
        for (let i = 0; i < sortedItems.length; i++) {
            if (i > 0 && sortedItems[i][scoreField] < sortedItems[i - 1][scoreField]) {
                rank = i + 1;
            }
            // Find the original item in the unsorted array and assign the rank
            const originalItem = items.find(item => item === sortedItems[i]);
            if(originalItem) {
                originalItem[rankField] = rank;
            }
        }
    };


    // --- DOM ELEMENTS ---
    const importFromFileBtn = $("importFromFileBtn"), gradeFile = $("gradeFile");
    const currentSemesterText = $("currentSemesterText"), examListBody = $("examListBody");
    const examAnalysisSelect = $("examAnalysisSelect"), subjectAnalysisSelect = $("subjectAnalysisSelect"), subjectAnalysisResult = $("subjectAnalysisResult");

    let currentSemester = "";
    
    // --- CORE INITIALIZATION ---
    const initialize = async () => {
        await initDB();
        const urlParams = new URLSearchParams(window.location.search);
        const appState = await getAppState();
        currentSemester = urlParams.get('semester') || appState.currentSemester;
        currentSemesterText.textContent = currentSemester;
        await renderExamList();
        await populateExamSelectorForAnalysis();
    };

    // --- RENDERING FUNCTIONS ---
    async function renderExamList() {
        examListBody.innerHTML = '';
        const allExams = await getAllExams();
        const semesterExams = allExams.filter(e => e.semester === currentSemester).sort((a,b) => new Date(b.date) - new Date(a.date));
        
        if (semesterExams.length === 0) {
            examListBody.innerHTML = `<tr><td colspan="4" style="text-align:center;">本学期暂无考试记录</td></tr>`;
            return;
        }
        semesterExams.forEach(exam => {
            const row = examListBody.insertRow();
            row.innerHTML = `<td>${exam.name}</td><td>${exam.date}</td><td>${(exam.subjects || []).length}</td><td><button class="secondary" onclick="showFullExamResults('${exam.id}')">查看成绩</button> <button class="danger" onclick="handleDeleteExam('${exam.id}')">删除考试</button></td>`;
        });
    }
    
    async function populateExamSelectorForAnalysis() {
        examAnalysisSelect.innerHTML = '<option value="">--请先选择一场考试--</option>';
        const allExams = await getAllExams();
        const semesterExams = allExams.filter(e => e.semester === currentSemester).sort((a,b) => new Date(b.date) - new Date(a.date));
        semesterExams.forEach(exam => {
            examAnalysisSelect.add(new Option(exam.name, exam.id));
        });
        // Reset subject dropdown as well
        populateSubjectSelectorForAnalysis(null);
    }

    async function populateSubjectSelectorForAnalysis(examId) {
        subjectAnalysisSelect.innerHTML = '<option value="">--再选择一个学科--</option>';
        subjectAnalysisResult.innerHTML = ''; // Clear previous results
        if (!examId) return;

        const allExams = await getAllExams();
        const exam = allExams.find(e => e.id === examId);
        if (!exam) return;

        (exam.subjects || []).forEach(s => subjectAnalysisSelect.add(new Option(s.name, s.name)));
    }


    // --- GLOBAL EVENT HANDLERS (attached to window) ---
    window.handleDeleteExam = async (examId) => {
        const allExams = await getAllExams();
        const exam = allExams.find(e => e.id === examId);
        if (!exam) return;
        
        if (confirm(`确定要删除考试 "${exam.name}" 吗？\n这将同时删除所有学生本次考试的成绩，此操作不可撤销！`)) {
            await deleteExam(examId);
            const allStudents = await getAllStudents();
            for (const student of allStudents) {
                const gradeCount = (student.grades || []).length;
                student.grades = (student.grades || []).filter(g => g.examId !== examId);
                if (student.grades.length < gradeCount) {
                     await putStudent(student);
                }
            }
            await renderExamList();
            await populateExamSelectorForAnalysis();
            showNotification(`考试 "${exam.name}" 已被成功删除。`, "success");
        }
    };

    window.showFullExamResults = async (examId) => {
        const allExams = await getAllExams();
        const exam = allExams.find(e => e.id === examId);
        if (!exam) return;
        
        $('examResultsTitle').textContent = `${exam.name} - 完整成绩单`;
        const container = $('examResultsTableContainer');
        
        let tableHTML = `<table style="width:100%;"><thead><tr><th>总排名</th><th>姓名</th>`;
        (exam.subjects || []).forEach(s => tableHTML += `<th>${s.name} (排名)</th>`);
        tableHTML += `<th>总分</th></tr></thead><tbody>`;
        
        const allStudents = await getAllStudents();
        const examGrades = allStudents.map(student => {
            const grade = (student.grades || []).find(g => g.examId === examId);
            return grade ? { studentName: student.name, grade } : null;
        }).filter(Boolean);

        examGrades.sort((a, b) => (a.grade.rank || 999) - (b.grade.rank || 999));

        examGrades.forEach(({studentName, grade}) => {
            tableHTML += `<tr><td>${grade.rank || 'N/A'}</td><td>${studentName}</td>`;
            (exam.subjects || []).forEach(s => {
                const subjectScore = (grade.subjects || []).find(gs => gs.name === s.name);
                tableHTML += `<td>${subjectScore ? `${subjectScore.score} (${subjectScore.rank || 'N/A'})` : '-'}</td>`;
            });
            const totalScore = grade.explicitTotalScore ?? (grade.subjects || []).reduce((sum, s) => sum + (Number(s.score) || 0), 0);
            tableHTML += `<td>${totalScore.toFixed(1)}</td></tr>`;
        });

        tableHTML += `</tbody></table>`;
        container.innerHTML = tableHTML;
        $('examResultsModal').classList.add('show');
    };

    // --- IMPORT LOGIC ---
    async function handleGradeFile(file) {
        if (!file) return showNotification("请先选择一个文件", "error");
        
        const examName = file.name.split('.').slice(0, -1).join('.');
        const allExams = await getAllExams();
        const existingExam = allExams.find(e => e.name === examName && e.semester === currentSemester);

        if (existingExam) {
            const modal = $('overwriteConfirmModal');
            $('overwriteExamName').textContent = examName;
            modal.classList.add('show');
            $('confirmOverwriteBtn').onclick = () => {
                modal.classList.remove('show');
                processImport(file, existingExam.id);
            };
        } else {
            processImport(file, null);
        }
    }

    function processImport(file, existingExamId = null) {
        const examName = file.name.split('.').slice(0, -1).join('.');
        const reader = new FileReader();
        reader.readAsArrayBuffer(file);
        reader.onload = async (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'array' });
                const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
                
                if (jsonData.length < 2) throw new Error("文件数据行数过少，无法解析。");
                const headerRow1 = jsonData[0].map(h => String(h || '').trim());
                const headerRow2 = jsonData[1].map(h => String(h || '').trim());
                let studentDataRowStartIndex = 2;

                const secondRowIsSubHeader = headerRow2.some(h => isNaN(parseFloat(h))) || headerRow2.length < headerRow1.length;
                if (!secondRowIsSubHeader) {
                    studentDataRowStartIndex = 1;
                }

                const columnMap = {};
                const subjects = [];

                for (let i = 0; i < headerRow1.length; i++) {
                    const h1 = headerRow1[i];
                    if (!h1) continue;

                    if (h1.includes("姓名")) { columnMap.nameIndex = i; continue; }
                    if (h1.includes("学号")) { columnMap.idIndex = i; continue; }

                    if (h1.includes("总分")) { columnMap.totalScoreIndex = i; }
                    else if (h1.includes("排名") || h1.includes("名次")) { columnMap.totalRankIndex = i; }
                    else {
                        const subject = { name: h1, scoreIndex: -1, rankIndex: -1 };
                        
                        if (secondRowIsSubHeader && headerRow1[i+1] === "") { // Merged cell
                            const h2_current = headerRow2[i];
                            const h2_next = headerRow2[i+1];
                            if(h2_current.includes("分")) subject.scoreIndex = i;
                            if(h2_current.includes("名")) subject.rankIndex = i;
                            if(h2_next && h2_next.includes("分")) subject.scoreIndex = i + 1;
                            if(h2_next && h2_next.includes("名")) subject.rankIndex = i + 1;
                            i++;
                        } else {
                            subject.scoreIndex = i;
                        }
                        subjects.push(subject);
                    }
                }

                if (columnMap.nameIndex === undefined) throw new Error("文件中必须包含'姓名'列。");

                const records = jsonData.slice(studentDataRowStartIndex);
                const extractedData = records.map(rec => {
                    const name = rec[columnMap.nameIndex];
                    if (!name) return null;

                    const studentSubjects = subjects.map(s => ({
                        name: s.name,
                        score: s.scoreIndex > -1 ? parseFloat(rec[s.scoreIndex]) : null,
                        rank: s.rankIndex > -1 ? parseInt(rec[s.rankIndex]) : null
                    })).filter(s => s.score !== null && !isNaN(s.score));

                    return {
                        name: String(name).trim(),
                        id: columnMap.idIndex > -1 ? String(rec[columnMap.idIndex]).trim() : null,
                        explicitTotalScore: columnMap.totalScoreIndex > -1 ? parseFloat(rec[columnMap.totalScoreIndex]) : null,
                        explicitTotalRank: columnMap.totalRankIndex > -1 ? parseInt(rec[columnMap.totalRankIndex]) : null,
                        subjects: studentSubjects,
                        totalScore: 0,
                        totalRank: null
                    };
                }).filter(Boolean);

                const needsTotalScoreCalc = extractedData.some(d => d.explicitTotalScore === null || isNaN(d.explicitTotalScore));
                if (needsTotalScoreCalc) {
                    extractedData.forEach(d => {
                        d.totalScore = d.subjects.reduce((sum, s) => sum + (s.score || 0), 0);
                    });
                } else {
                    extractedData.forEach(d => { d.totalScore = d.explicitTotalScore; });
                }

                if (extractedData.some(d => d.explicitTotalRank === null || isNaN(d.explicitTotalRank))) {
                    assignRanks(extractedData, 'totalScore', 'totalRank');
                } else {
                    extractedData.forEach(d => { d.totalRank = d.explicitTotalRank; });
                }

                subjects.forEach(subject => {
                    if (extractedData.some(d => { const sub = d.subjects.find(s => s.name === subject.name); return sub && (sub.rank === null || isNaN(sub.rank)); })) {
                        const subjectData = extractedData.map(d => d.subjects.find(s => s.name === subject.name)).filter(Boolean);
                        assignRanks(subjectData, 'score', 'rank');
                    }
                });

                const newExamId = existingExamId || Date.now().toString();
                const newExam = { 
                    id: newExamId, 
                    name: examName, 
                    date: new Date().toISOString().split('T')[0], 
                    semester: currentSemester, 
                    subjects: subjects.map(s => ({ name: s.name, fullScore: 100 }))
                };
                
                let allStudents = await getAllStudents();
                let updatedCount = 0;
                
                for (const studentData of extractedData) {
                    let student = allStudents.find(s => s.name === studentData.name);
                    if (!student) {
                        const newId = studentData.id || `S${Date.now()}${Math.random().toString(36).substr(2, 5)}`;
                        student = { id: newId, name: studentData.name, gender: '未指定', grades: [], disciplineRecords: [] };
                        allStudents.push(student);
                    }
                    
                    const studentGrade = {
                        examId: newExam.id, examName: newExam.name, date: newExam.date,
                        semester: newExam.semester, 
                        rank: studentData.totalRank,
                        explicitTotalScore: studentData.totalScore,
                        subjects: studentData.subjects
                    };
                    
                    student.grades = (student.grades || []).filter(g => g.examId !== newExam.id);
                    student.grades.push(studentGrade);
                    await putStudent(student);
                    updatedCount++;
                }

                if(existingExamId) await deleteExam(existingExamId);
                await putExam(newExam);

                await renderExamList();
                await populateExamSelectorForAnalysis();
                const actionText = existingExamId ? '覆盖' : '导入';
                showNotification(`${actionText}成功！考试'${examName}'已处理，更新了${updatedCount}名学生的成绩。`, "success");
                gradeFile.value = '';

            } catch (error) { 
                console.error("Import failed:", error);
                showNotification(`处理失败: ${error.message}`, "error"); 
            }
        };
        reader.onerror = () => showNotification("读取文件失败", "error");
    }

    // --- EVENT LISTENERS ---
    importFromFileBtn.onclick = () => handleGradeFile(gradeFile.files[0]);
    
    examAnalysisSelect.onchange = () => {
        populateSubjectSelectorForAnalysis(examAnalysisSelect.value);
    };

    subjectAnalysisSelect.onchange = async () => {
        const selectedExamId = examAnalysisSelect.value;
        const subjectName = subjectAnalysisSelect.value;
        subjectAnalysisResult.innerHTML = '';
        if (!selectedExamId || !subjectName) return;

        const examName = examAnalysisSelect.options[examAnalysisSelect.selectedIndex].text;
        let tableHTML = `<h3>"${subjectName}" 在 "${examName}" 中的成绩排名</h3><table><thead><tr><th>排名</th><th>学生</th><th>分数</th></tr></thead><tbody>`;
        
        const allStudents = await getAllStudents();
        const subjectScores = [];
        allStudents.forEach(student => {
            const grade = (student.grades || []).find(g => g.examId === selectedExamId);
            if (grade) {
                const subjectScore = (grade.subjects || []).find(s => s.name === subjectName);
                if (subjectScore) {
                    subjectScores.push({ studentName: student.name, score: Number(subjectScore.score) });
                }
            }
        });

        subjectScores.sort((a,b) => b.score - a.score);
        subjectScores.forEach((item, index) => {
            let rank = index + 1;
            if (index > 0 && subjectScores[index].score === subjectScores[index - 1].score) {
                // Find the rank of the previous student in the table to handle ties
                const prevRow = subjectAnalysisResult.querySelector(`tr:nth-child(${index})`);
                if(prevRow) rank = parseInt(prevRow.cells[0].textContent);
            }
            tableHTML += `<tr><td>${rank}</td><td>${item.studentName}</td><td>${item.score}</td></tr>`;
        });

        tableHTML += `</tbody></table>`;
        subjectAnalysisResult.innerHTML = tableHTML;
    };

    // --- PAGE START ---
    initialize();
});
