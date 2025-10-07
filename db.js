/**
 * 使用 IndexedDB 数据库核心 (V5.0)
 * 这是一个独立的模块，负责所有数据的存储和读取。
 * 它使用 idb 库来简化 IndexedDB 的操作。
 */

const DB_NAME = 'StudentDB';
const DB_VERSION = 1;
let db;

async function initDB() {
    if (db) return;
    db = await idb.openDB(DB_NAME, DB_VERSION, {
        upgrade(db) {
            if (!db.objectStoreNames.contains('students')) {
                db.createObjectStore('students', { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains('exams')) {
                db.createObjectStore('exams', { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains('semesters')) {
                db.createObjectStore('semesters');
            }
            if (!db.objectStoreNames.contains('appState')) {
                db.createObjectStore('appState');
            }
        },
    });
}

// --- 学生数据操作 ---
async function getAllStudents() {
    await initDB();
    return db.getAll('students');
}
async function putStudent(student) {
    await initDB();
    return db.put('students', student);
}
async function deleteStudent(studentId) {
    await initDB();
    return db.delete('students', studentId);
}

// --- 考试数据操作 ---
async function getAllExams() {
    await initDB();
    return db.getAll('exams');
}
async function putExam(exam) {
    await initDB();
    return db.put('exams', exam);
}
async function deleteExam(examId) {
    await initDB();
    return db.delete('exams', examId);
}

// --- 学期数据操作 ---
async function getAllSemesters() {
    await initDB();
    const keys = await db.getAllKeys('semesters');
    return keys.length > 0 ? keys : ["2024-2025-1"];
}
async function putSemester(semester) {
    await initDB();
    return db.put('semesters', semester, semester);
}

// --- 应用状态操作 ---
async function getAppState() {
    await initDB();
    const state = await db.get('appState', 'currentState');
    return state || { currentSemester: "2024-2025-1" };
}
async function saveAppState(state) {
    await initDB();
    return db.put('appState', state, 'currentState');
}

// --- 高级操作 ---
async function exportAllData() {
    await initDB();
    const students = await db.getAll('students');
    const exams = await db.getAll('exams');
    const semesters = await db.getAllKeys('semesters');
    const appState = await db.get('appState', 'currentState');

    return {
        students,
        exams,
        semesters,
        appState,
        exportDate: new Date().toISOString()
    };
}

async function importAllData(dataToImport) {
    await initDB();
    if (!dataToImport || !Array.isArray(dataToImport.students) || !Array.isArray(dataToImport.exams)) {
        throw new Error("无效的备份文件格式。");
    }
    
    await clearAllData();

    const importTx = db.transaction(['students', 'exams', 'semesters', 'appState'], 'readwrite');
    const promises = [];
    
    dataToImport.students.forEach(student => promises.push(importTx.objectStore('students').put(student)));
    dataToImport.exams.forEach(exam => promises.push(importTx.objectStore('exams').put(exam)));
    (dataToImport.semesters || []).forEach(semester => promises.push(importTx.objectStore('semesters').put(semester, semester)));
    if (dataToImport.appState) {
        promises.push(importTx.objectStore('appState').put(dataToImport.appState, 'currentState'));
    }

    await Promise.all(promises);
    await importTx.done;
}

async function clearAllData() {
    await initDB();
    await db.clear('students');
    await db.clear('exams');
    await db.clear('semesters');
    await db.clear('appState');
}

