class GanttChart {
    constructor() {
        this.workPackages = [];
        this.currentDate = new Date(2026, 1, 1); // 2026年2月1日
        this.dragData = null;
        this.draggedWorkPackageId = null; // ドラッグ中の工程ID
        this.draggedTaskId = null; // ドラッグ中のタスクID
        this.draggedTaskSourceWpId = null; // ドラッグ中のタスクの元の工程ID
        this.maxWorkPackages = 5000;
        this.daysInView = 365; // 12ヶ月分を表示
        this.monthsInView = 12; // 12ヶ月表示
        this.showTasks = true; // タスク一覧の表示非表示フラグ
        this.addWorkPackageMode = false; // 工程追加モード
        this.currentLanguage = 'ja'; // 現在の言語（デフォルト：日本語）
        this.zoomLevel = 100; // ズームレベル（100%から400%まで）
        this.pixelPerDay = 30; // 1日あたりのピクセル数（ズームに応じて変わる）
        this.unitType = 'day'; // 表示単位：'day', 'week', 'month'
        this.addTaskMode = false; // タスク追加モード
        this.editMode = false; // 編集モード（タッチ操作用）
        this._rafId = null; // ドラッグ時のrequestAnimationFrame ID
        this.translations = {
            ja: {
                title: 'Groops',
                addTask: '工程を追加',
                addWorkPackageMode: '工程を挿入',
                addWorkPackageModeActive: '工程挿入モード（クリックで追加）',
                addTaskMode: 'タスクを追加',
                addTaskModeActive: 'タスク追加モード（クリックで追加）',
                tasksHide: '▼ タスク非表示',
                tasksShow: '▶ タスク表示',
                reset: 'リセット',
                prevMonth: '◀ 前月',
                nextMonth: '次月 ▶',
                workPackages: '工程',
                addBtn: '+',
                deleteBtn: '削除',
                confirmReset: 'すべてのデータをリセットしてもいいですか？',
                newWorkPackage: '新しい工程',
                newTask: '新しいタスク',
                maxWorkPackagesMsg: '個の工程までしか追加できません最大',
                englishLabel: 'English',
                unitDay: '日',
                unitWeek: '週',
                unitMonth: '月',
                exportToExcel: 'エクセルにエクスポート',
                editMode: '✏️ 編集モード',
                editModeActive: '✏️ 編集中（タップで解除）'
            },
            en: {
                title: 'Groops',
                addTask: 'Add Work Package',
                addWorkPackageMode: 'Insert Work Package',
                addWorkPackageModeActive: 'Insert WP Mode (Click to add)',
                addTaskMode: 'Add Task',
                addTaskModeActive: 'Add Task Mode (Click to add)',
                tasksHide: '▼ Hide Tasks',
                tasksShow: '▶ Show Tasks',
                reset: 'Reset',
                prevMonth: '◀ Prev',
                nextMonth: 'Next ▶',
                workPackages: 'Work Packages',
                addBtn: '+',
                deleteBtn: 'Delete',
                confirmReset: 'Reset all data?',
                newWorkPackage: 'New Work Package',
                newTask: 'New Task',
                maxWorkPackagesMsg: 'Can only add up to',
                englishLabel: '日本語',
                unitDay: 'Day',
                unitWeek: 'Week',
                unitMonth: 'Month',
                exportToExcel: 'Export to Excel',
                editMode: '✏️ Edit Mode',
                editModeActive: '✏️ Editing (tap to exit)'
            }
        };
        this.init();
    }

    t(key) {
        return this.translations[this.currentLanguage][key] || key;
    }

    loadLanguage() {
        const saved = localStorage.getItem('ganttLanguage');
        if (saved && (saved === 'ja' || saved === 'en')) {
            this.currentLanguage = saved;
        }
    }

    saveLanguage() {
        localStorage.setItem('ganttLanguage', this.currentLanguage);
    }

    toggleLanguage() {
        this.currentLanguage = this.currentLanguage === 'ja' ? 'en' : 'ja';
        this.saveLanguage();
        this.render();
    }

    toggleAddWorkPackageMode() {
        this.addWorkPackageMode = !this.addWorkPackageMode;
        const btn = document.getElementById('addTaskBtn');
        const taskList = document.getElementById('taskList');
        
        if (this.addWorkPackageMode) {
            btn.textContent = this.t('addWorkPackageModeActive');
            btn.classList.add('btn-active-mode');
            taskList.classList.add('add-workpackage-cursor');
        } else {
            btn.textContent = this.t('addWorkPackageMode');
            btn.classList.remove('btn-active-mode');
            taskList.classList.remove('add-workpackage-cursor');
        }
        this.render();
    }

    toggleAddTaskMode() {
        this.addTaskMode = !this.addTaskMode;
        const btn = document.getElementById('addTaskModeBtn');
        const ganttContent = document.getElementById('ganttContent');
        
        if (this.addTaskMode) {
            btn.textContent = this.t('addTaskModeActive');
            btn.classList.add('btn-active-mode');
            ganttContent.classList.add('add-task-cursor');
        } else {
            btn.textContent = this.t('addTaskMode');
            btn.classList.remove('btn-active-mode');
            ganttContent.classList.remove('add-task-cursor');
        }
    }

    toggleEditMode() {
        this.editMode = !this.editMode;
        const btn = document.getElementById('editModeBtn');
        const container = document.querySelector('.gantt-container');

        if (this.editMode) {
            btn.textContent = this.t('editModeActive');
            btn.classList.add('active');
            container.classList.add('edit-mode-active');
        } else {
            btn.textContent = this.t('editMode');
            btn.classList.remove('active');
            container.classList.remove('edit-mode-active');
        }
    }

    zoomIn() {
        if (this.zoomLevel < 400) {
            this.zoomLevel += 25;
            this.updateZoom();
        }
    }

    zoomOut() {
        if (this.zoomLevel > 25) {
            this.zoomLevel -= 25;
            this.updateZoom();
        }
    }

    updateZoom() {
        // ズームに応じてピクセル数を更新
        const unitDays = this.getUnitDays();
        this.pixelPerDay = (30 * this.zoomLevel / 100) / unitDays;
        document.getElementById('zoomLevel').textContent = `${this.zoomLevel}%`;
        this.render();
    }

    setUnit(unit) {
        if (['day', 'week', 'month'].includes(unit)) {
            this.unitType = unit;
            this.updateUnitButtons();
            // ズームレベルをリセット
            this.zoomLevel = 100;
            this.updateZoom();
        }
    }

    getUnitDays() {
        switch (this.unitType) {
            case 'week': return 7;
            case 'month': return 30;
            default: return 1; // day
        }
    }

    getUnitLabel(index) {
        const startDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1);
        const date = new Date(startDate);

        switch (this.unitType) {
            case 'week':
                // 週番号を表示
                const weekStart = new Date(date);
                weekStart.setDate(weekStart.getDate() + index * 7);
                return `W${Math.ceil((weekStart.getDate()) / 7)}`;
            case 'month':
                // 月を表示
                const monthDate = new Date(date);
                monthDate.setMonth(monthDate.getMonth() + index);
                if (monthDate.getMonth() === 0) {
                    return `${monthDate.getFullYear()}年1月`.substring(5);
                }
                return `${monthDate.getMonth() + 1}月`;
            default: // day
                const dayDate = new Date(date);
                dayDate.setDate(dayDate.getDate() + index);
                if (dayDate.getDate() === 1) {
                    return `${dayDate.getMonth() + 1}月`;
                }
                return `${dayDate.getDate()}`;
        }
    }

    updateUnitButtons() {
        const buttons = {
            'day': document.getElementById('unitDay'),
            'week': document.getElementById('unitWeek'),
            'month': document.getElementById('unitMonth')
        };

        Object.keys(buttons).forEach(unit => {
            if (unit === this.unitType) {
                buttons[unit].classList.add('btn-active');
            } else {
                buttons[unit].classList.remove('btn-active');
            }
        });
    }

    init() {
        this.loadLanguage();
        this.loadData();
        this.updateZoom(); // 初期ズームを設定
        this.setupEventListeners();
        this.render();
    }

    setupEventListeners() {
        document.getElementById('addTaskBtn').addEventListener('click', () => this.toggleAddWorkPackageMode());
        document.getElementById('addTaskModeBtn').addEventListener('click', () => this.toggleAddTaskMode());
        document.getElementById('toggleTasksBtn').addEventListener('click', () => this.toggleTasksVisibility());
        document.getElementById('resetBtn').addEventListener('click', () => this.resetData());
        document.getElementById('prevMonth').addEventListener('click', () => this.prevMonth());
        document.getElementById('nextMonth').addEventListener('click', () => this.nextMonth());
        document.getElementById('langToggleBtn').addEventListener('click', () => this.toggleLanguage());
        document.getElementById('exportBtn').addEventListener('click', () => this.exportToExcel());
        document.getElementById('editModeBtn').addEventListener('click', () => this.toggleEditMode());
        
        // ズームコントロール
        document.getElementById('zoomOutBtn').addEventListener('click', () => this.zoomOut());
        document.getElementById('zoomInBtn').addEventListener('click', () => this.zoomIn());
        
        // ユニット切り替え
        document.getElementById('unitDay').addEventListener('click', () => this.setUnit('day'));
        document.getElementById('unitWeek').addEventListener('click', () => this.setUnit('week'));
        document.getElementById('unitMonth').addEventListener('click', () => this.setUnit('month'));
        
        document.getElementById('ganttContent').addEventListener('mousedown', (e) => this.onMouseDown(e));
        document.addEventListener('mousemove', (e) => this.onMouseMove(e));
        document.addEventListener('mouseup', (e) => this.onMouseUp(e));

        // タッチイベント（スマホ対応：長押しでバーをドラッグ）
        document.getElementById('ganttContent').addEventListener('touchstart', (e) => this.onTouchStart(e), { passive: false });
        document.addEventListener('touchmove', (e) => this.onTouchMove(e), { passive: false });
        document.addEventListener('touchend', (e) => this.onTouchEnd(e));
        
        // スクロール同期
        const ganttSidebar = document.querySelector('.gantt-sidebar');
        const ganttContent = document.getElementById('ganttContent');
        const ganttHeader = document.getElementById('ganttHeader');
        

        
        // ガントコンテンツ（右パネル）のスクロール - 水平と垂直の両方を同期
        ganttContent.addEventListener('scroll', (e) => {
            // 水平スクロール：日付ヘッダーと同期
            ganttHeader.scrollLeft = e.target.scrollLeft;
            
            // 垂直スクロール：左パネルと同期（無限ループを防ぐ）
            if (!this._syncing && ganttSidebar) {
                this._syncing = true;
                ganttSidebar.scrollTop = e.target.scrollTop;

                this._syncing = false;
            }
        });
        
        // ガントサイドバー（左パネル）のスクロール - 垂直スクロールを右パネルと同期
        if (ganttSidebar) {
            ganttSidebar.addEventListener('scroll', (e) => {
                if (!this._syncing) {
                    this._syncing = true;
                    ganttContent.scrollTop = e.target.scrollTop;

                    this._syncing = false;
                }
            });
        }
    }

    loadData() {
        const saved = localStorage.getItem('ganttData');

        
        if (saved) {
            const data = JSON.parse(saved);

            this.workPackages = data.map(wp => ({
                ...wp,
                tasks: wp.tasks.map(task => ({
                    ...task,
                    startDate: new Date(task.startDate)
                }))
            }));
        } else {

            // デフォルトデータ
            this.workPackages = [
                {
                    id: 1,
                    name: '企画',
                    tasks: [
                        { id: 101, name: 'プロジェクト計画', startDate: new Date(2026, 1, 1), duration: 3 },
                        { id: 102, name: '要件定義', startDate: new Date(2026, 1, 5), duration: 5 }
                    ]
                },
                {
                    id: 2,
                    name: '設計',
                    tasks: [
                        { id: 201, name: '基本設計', startDate: new Date(2026, 1, 10), duration: 5 },
                        { id: 202, name: '詳細設計', startDate: new Date(2026, 1, 16), duration: 5 }
                    ]
                },
                {
                    id: 3,
                    name: '開発',
                    tasks: [
                        { id: 301, name: 'バックエンド開発', startDate: new Date(2026, 1, 21), duration: 10 },
                        { id: 302, name: 'フロントエンド開発', startDate: new Date(2026, 1, 22), duration: 10 }
                    ]
                },
                {
                    id: 4,
                    name: 'テスト',
                    tasks: [
                        { id: 401, name: '単体テスト', startDate: new Date(2026, 2, 3), duration: 3 },
                        { id: 402, name: '統合テスト', startDate: new Date(2026, 2, 7), duration: 3 }
                    ]
                }
            ];

        }
    }

    saveData() {
        const data = this.workPackages.map(wp => ({
            ...wp,
            tasks: wp.tasks.map(task => ({
                ...task,
                startDate: task.startDate.toISOString()
            }))
        }));
        localStorage.setItem('ganttData', JSON.stringify(data));
    }

    addWorkPackageAtIndex(index) {
        const newId = Math.max(0, ...this.workPackages.map(wp => wp.id)) + 1;
        const newWorkPackage = {
            id: newId,
            name: `${this.t('newWorkPackage')}${newId}`,
            tasks: [
                {
                    id: newId * 1000 + 1,
                    name: `${this.t('newTask')}${newId * 1000 + 1}`,
                    startDate: new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1),
                    duration: 7
                }
            ]
        };
        
        this.workPackages.splice(index, 0, newWorkPackage);
        this.saveData();
        this.addWorkPackageMode = false;
        this.render();
    }

    addWorkPackage() {
        if (this.workPackages.length >= this.maxWorkPackages) {
            alert(`${this.t('maxWorkPackagesMsg')}${this.maxWorkPackages}${this.currentLanguage === 'ja' ? '個の工程までしか追加できません' : 'work packages.'}`);
            return;
        }
        const newId = Math.max(0, ...this.workPackages.map(wp => wp.id)) + 1;
        const newWorkPackage = {
            id: newId,
            name: `${this.t('newWorkPackage')}${newId}`,
            tasks: [
                { id: newId * 100 + 1, name: this.t('newTask'), startDate: new Date(this.currentDate), duration: 5 }
            ]
        };
        this.workPackages.push(newWorkPackage);

        this.saveData();
        this.render();
    }

    resetData() {
        if (confirm(this.t('confirmReset'))) {

            localStorage.removeItem('ganttData');
            this.workPackages = [];
            this.loadData();
            this.render();

        }
    }

    toggleTasksVisibility() {
        this.showTasks = !this.showTasks;
        const btn = document.getElementById('toggleTasksBtn');
        
        if (this.showTasks) {
            btn.textContent = this.t('tasksHide');
        } else {
            btn.textContent = this.t('tasksShow');
        }

        // 全ての .tasks-list の表示/非表示を切り替える
        const tasksLists = document.querySelectorAll('.tasks-list');
        tasksLists.forEach(tasksList => {
            tasksList.style.display = this.showTasks ? 'block' : 'none';
        });


        
        // 表示状態に合わせてガント側もリレンダリング
        this.render();
    }

    deleteWorkPackage(id) {
        this.workPackages = this.workPackages.filter(wp => wp.id !== id);
        this.saveData();
        this.render();
    }

    addTaskToWorkPackage(wpId) {
        const wp = this.workPackages.find(w => w.id === wpId);
        if (!wp) return;
        
        const newId = Math.max(0, ...wp.tasks.map(t => t.id)) + 1;
        const newTask = {
            id: newId,
            name: '新しいタスク',
            startDate: new Date(this.currentDate),
            duration: 5
        };
        wp.tasks.push(newTask);
        this.saveData();
        this.render();
    }

    deleteTask(taskId) {
        this.workPackages.forEach(wp => {
            wp.tasks = wp.tasks.filter(t => t.id !== taskId);
        });
        this.saveData();
        this.render();
    }

    toggleTaskMenu(taskEl, taskId, wpId) {
        // 既存のメニューを削除
        document.querySelectorAll('.task-dropdown-menu').forEach(menu => menu.remove());
        
        // メニューを作成
        const menu = document.createElement('div');
        menu.className = 'task-dropdown-menu';
        
        // 削除
        const deleteItem = document.createElement('div');
        deleteItem.className = 'task-menu-item';
        deleteItem.innerHTML = '<span>🗑️</span> ' + (this.currentLanguage === 'ja' ? '削除' : 'Delete');
        deleteItem.onclick = () => {
            this.deleteTask(taskId);
            menu.remove();
        };
        
        // 上に移動
        const moveUpItem = document.createElement('div');
        moveUpItem.className = 'task-menu-item';
        moveUpItem.innerHTML = '<span>⬆️</span> ' + (this.currentLanguage === 'ja' ? '上に移動' : 'Move Up');
        moveUpItem.onclick = () => {
            this.moveTaskUp(taskId, wpId);
            menu.remove();
        };
        
        // 下に移動
        const moveDownItem = document.createElement('div');
        moveDownItem.className = 'task-menu-item';
        moveDownItem.innerHTML = '<span>⬇️</span> ' + (this.currentLanguage === 'ja' ? '下に移動' : 'Move Down');
        moveDownItem.onclick = () => {
            this.moveTaskDown(taskId, wpId);
            menu.remove();
        };
        
        menu.appendChild(deleteItem);
        menu.appendChild(moveUpItem);
        menu.appendChild(moveDownItem);
        
        taskEl.appendChild(menu);
        
        // メニューの外側をクリックしたら閉じる
        setTimeout(() => {
            document.addEventListener('click', function closeMenu(e) {
                if (!menu.contains(e.target)) {
                    menu.remove();
                    document.removeEventListener('click', closeMenu);
                }
            });
        }, 0);
    }

    moveTaskUp(taskId, wpId) {
        const wp = this.workPackages.find(w => w.id === wpId);
        if (!wp) return;
        
        const index = wp.tasks.findIndex(t => t.id === taskId);
        if (index > 0) {
            [wp.tasks[index - 1], wp.tasks[index]] = [wp.tasks[index], wp.tasks[index - 1]];
            this.saveData();
            this.render();
        }
    }

    moveTaskDown(taskId, wpId) {
        const wp = this.workPackages.find(w => w.id === wpId);
        if (!wp) return;
        
        const index = wp.tasks.findIndex(t => t.id === taskId);
        if (index >= 0 && index < wp.tasks.length - 1) {
            [wp.tasks[index], wp.tasks[index + 1]] = [wp.tasks[index + 1], wp.tasks[index]];
            this.saveData();
            this.render();
        }
    }

    updateWorkPackageName(wpId, newName) {
        const wp = this.workPackages.find(w => w.id === wpId);
        if (wp && newName.trim()) {
            wp.name = newName.trim();
            this.saveData();
            this.render();
        }
    }

    updateTaskName(taskId, newName) {
        let found = false;
        this.workPackages.forEach(wp => {
            const task = wp.tasks.find(t => t.id === taskId);
            if (task && newName.trim()) {
                task.name = newName.trim();
                found = true;
            }
        });
        if (found) {
            this.saveData();
            this.render();
        }
    }

    startEditingWorkPackageName(element, wpId, currentName) {
        const input = document.createElement('input');
        input.type = 'text';
        input.value = currentName;
        input.style.width = '100%';
        input.style.padding = '2px';
        input.style.border = '1px solid #3498db';
        input.style.borderRadius = '2px';
        input.style.fontSize = '11px';

        const finishEdit = () => {
            const newName = input.value.trim();
            if (newName && newName !== currentName) {
                this.updateWorkPackageName(wpId, newName);
            } else {
                this.render();
            }
        };

        input.onkeydown = (e) => {
            if (e.key === 'Enter') finishEdit();
            if (e.key === 'Escape') this.render();
        };
        input.onblur = finishEdit;

        element.parentElement.replaceChild(input, element);
        input.focus();
        input.select();
    }

    startEditingTaskName(element, taskId, currentName) {
        const input = document.createElement('input');
        input.type = 'text';
        input.value = currentName;
        input.style.width = '100%';
        input.style.padding = '2px';
        input.style.border = '1px solid #3498db';
        input.style.borderRadius = '2px';
        input.style.fontSize = '10px';

        const finishEdit = () => {
            const newName = input.value.trim();
            if (newName && newName !== currentName) {
                this.updateTaskName(taskId, newName);
            } else {
                this.render();
            }
        };

        input.onkeydown = (e) => {
            if (e.key === 'Enter') finishEdit();
            if (e.key === 'Escape') this.render();
        };
        input.onblur = finishEdit;

        element.parentElement.replaceChild(input, element);
        input.focus();
        input.select();
    }

    prevMonth() {
        this.currentDate.setMonth(this.currentDate.getMonth() - 1);
        this.render();
    }

    nextMonth() {
        this.currentDate.setMonth(this.currentDate.getMonth() + 1);
        this.render();
    }

    render() {
        this.updateUILanguage();
        this.renderHeader();
        this.renderTasks();
        this.renderGantt();
        // レイアウト計算完了後に高さを同期
        requestAnimationFrame(() => this.syncRowHeights());
    }

    updateUILanguage() {
        // ボタンテキストを更新
        document.getElementById('addTaskBtn').textContent = this.t('addTask');
        const addTaskModeBtn = document.getElementById('addTaskModeBtn');
        addTaskModeBtn.textContent = this.addTaskMode ? this.t('addTaskModeActive') : this.t('addTaskMode');
        document.getElementById('toggleTasksBtn').textContent = this.showTasks ? this.t('tasksHide') : this.t('tasksShow');
        document.getElementById('exportBtn').textContent = this.t('exportToExcel');
        document.getElementById('resetBtn').textContent = this.t('reset');
        document.getElementById('prevMonth').textContent = this.t('prevMonth');
        document.getElementById('nextMonth').textContent = this.t('nextMonth');
        document.getElementById('langToggleBtn').textContent = this.t('englishLabel');
        
        // ユニットボタンのテキストを更新
        document.getElementById('unitDay').textContent = this.t('unitDay');
        document.getElementById('unitWeek').textContent = this.t('unitWeek');
        document.getElementById('unitMonth').textContent = this.t('unitMonth');
        
        // ズームレベル表示を更新
        document.getElementById('zoomLevel').textContent = `${this.zoomLevel}%`;

        // 編集モードボタンのテキストを更新
        const editModeBtn = document.getElementById('editModeBtn');
        editModeBtn.textContent = this.editMode ? this.t('editModeActive') : this.t('editMode');
    }

    syncRowHeights() {
        // 高さはCSSで固定するため、この関数は無効化

        return;
        
        const ganttRows = document.querySelectorAll('.gantt-row');
        
        if (this.showTasks) {
            // タスク表示オン：工程ヘッダー行とタスク行を順序通りに対応させる
            let ganttRowIndex = 0;
            const wpItems = document.querySelectorAll('.workpackage-item');
            
            wpItems.forEach((wpItem) => {
                // 1. 工程ヘッダー行
                const wpHeader = wpItem.querySelector('.workpackage-header');
                if (wpHeader && ganttRows[ganttRowIndex]) {
                    const height = wpHeader.offsetHeight;
                    ganttRows[ganttRowIndex].style.height = height + 'px';
                    const container = ganttRows[ganttRowIndex].querySelector('.gantt-bar-container');
                    if (container) container.style.height = height + 'px';
                }
                ganttRowIndex++;
                
                // 2. タスク行をすべて対応させる
                const taskItems = wpItem.querySelectorAll('.task-item');
                taskItems.forEach((taskItem) => {
                    if (ganttRows[ganttRowIndex]) {
                        const height = taskItem.offsetHeight;
                        ganttRows[ganttRowIndex].style.height = height + 'px';
                        const container = ganttRows[ganttRowIndex].querySelector('.gantt-bar-container');
                        if (container) container.style.height = height + 'px';
                    }
                    ganttRowIndex++;
                });
            });
        } else {
            // タスク表示オフ：各工程全体に対応させる
            const wpItems = document.querySelectorAll('.workpackage-item');
            wpItems.forEach((item, index) => {
                if (ganttRows[index]) {
                    const height = item.offsetHeight;
                    ganttRows[index].style.height = height + 'px';
                    
                    const container = ganttRows[index].querySelector('.gantt-bar-container');
                    if (container) {
                        container.style.height = height + 'px';
                    }
                }
            });
        }
    }

    renderHeader() {
        const year = this.currentDate.getFullYear();
        const month = this.currentDate.getMonth();
        if (this.currentLanguage === 'ja') {
            document.getElementById('currentMonth').textContent = `${year}年${month + 1}月から12ヶ月表示`;
        } else {
            document.getElementById('currentMonth').textContent = `${year} - Month ${month + 1} (12 months)`;
        }

        const headerContainer = document.getElementById('ganttHeader');
        headerContainer.innerHTML = '';
        
        // ラッパーを作成して、これに日付要素を追加
        const wrapper = document.createElement('div');
        wrapper.style.display = 'flex';
        
        const unitDays = this.getUnitDays();
        const unitCount = Math.ceil(this.daysInView / unitDays);
        const totalWidth = unitCount * this.pixelPerDay * unitDays;
        wrapper.style.width = totalWidth + 'px';

        const startDate = new Date(year, month, 1);
        
        for (let i = 0; i < unitCount; i++) {
            const headerWidth = this.pixelPerDay * unitDays;
            const dayEl = document.createElement('div');
            dayEl.className = 'header-day';
            dayEl.style.width = headerWidth + 'px';
            dayEl.textContent = this.getUnitLabel(i);
            dayEl.style.fontWeight = 'bold';
            dayEl.style.backgroundColor = '#ecf0f1';
            wrapper.appendChild(dayEl);
        }
        headerContainer.appendChild(wrapper);
    }

    renderTasks() {
        const taskList = document.getElementById('taskList');
        taskList.innerHTML = '';

        // sidebar-headerを更新
        const sidebarHeader = document.querySelector('.sidebar-header');
        if (sidebarHeader) {
            sidebarHeader.textContent = this.t('workPackages');
        }



        this.workPackages.forEach((wp, wpIndex) => {

            
            const wpEl = document.createElement('div');
            wpEl.className = 'workpackage-item';
            wpEl.draggable = true;
            wpEl.dataset.wpId = wp.id;
            wpEl.dataset.dragType = 'workpackage';
            
            // 工程追加モード時のクリックイベント
            wpEl.addEventListener('click', (e) => {
                if (this.addWorkPackageMode && !e.target.classList.contains('delete-btn') && !e.target.classList.contains('add-task-btn')) {
                    this.addWorkPackageAtIndex(wpIndex);
                }
            });
            
            // 工程追加モード時のホバーエフェクト
            wpEl.addEventListener('mouseenter', () => {
                if (this.addWorkPackageMode) {
                    wpEl.style.backgroundColor = '#e3f2fd';
                    wpEl.style.cursor = 'pointer';
                }
            });
            
            wpEl.addEventListener('mouseleave', () => {
                if (this.addWorkPackageMode) {
                    wpEl.style.backgroundColor = '';
                    wpEl.style.cursor = '';
                }
            });
            
            // ドラッグイベント
            wpEl.addEventListener('dragstart', (e) => {
                if (this.addWorkPackageMode) return;
                this.onWorkPackageDragStart(e, wp.id);
            });
            wpEl.addEventListener('dragover', (e) => this.onWorkPackageDragOver(e));
            wpEl.addEventListener('drop', (e) => this.onWorkPackageDrop(e, wp.id));
            wpEl.addEventListener('dragend', (e) => this.onWorkPackageDragEnd(e));
            wpEl.addEventListener('dragleave', (e) => this.onWorkPackageDragLeave(e));
            
            const wpHeaderEl = document.createElement('div');
            wpHeaderEl.className = 'workpackage-header';
            
            const wpNameEl = document.createElement('span');
            wpNameEl.className = 'workpackage-name';
            wpNameEl.textContent = wp.name;
            wpNameEl.title = wp.name;
            wpNameEl.style.cursor = 'pointer';
            wpNameEl.onclick = () => this.startEditingWorkPackageName(wpNameEl, wp.id, wp.name);
            
            const wpAddBtn = document.createElement('button');
            wpAddBtn.className = 'add-task-btn';
            wpAddBtn.textContent = this.t('addBtn');
            wpAddBtn.onclick = () => this.addTaskToWorkPackage(wp.id);
            
            const wpDeleteBtn = document.createElement('button');
            wpDeleteBtn.className = 'delete-btn';
            wpDeleteBtn.textContent = this.t('deleteBtn');
            wpDeleteBtn.onclick = () => this.deleteWorkPackage(wp.id);
            
            wpHeaderEl.appendChild(wpNameEl);
            wpHeaderEl.appendChild(wpAddBtn);
            wpHeaderEl.appendChild(wpDeleteBtn);
            wpEl.appendChild(wpHeaderEl);
            
            // タスク一覧
            const tasksEl = document.createElement('div');
            tasksEl.className = 'tasks-list';
            tasksEl.dataset.wpId = wp.id;
            if (!this.showTasks) {
                tasksEl.style.display = 'none';
            }
            
            // タスクコンテナにもドロップイベントを設定
            tasksEl.addEventListener('dragover', (e) => this.onTaskDragOver(e));
            tasksEl.addEventListener('drop', (e) => this.onTaskDrop(e, wp.id));
            tasksEl.addEventListener('dragleave', (e) => this.onTaskDragLeave(e));
            
            wp.tasks.forEach((task, taskIndex) => {

                
                const taskEl = document.createElement('div');
                taskEl.className = 'task-item';
                taskEl.draggable = true;
                taskEl.dataset.taskId = task.id;
                taskEl.dataset.wpId = wp.id;
                taskEl.dataset.dragType = 'task';
                
                // タスクドラッグイベント
                taskEl.addEventListener('dragstart', (e) => this.onTaskDragStart(e, task.id, wp.id));
                taskEl.addEventListener('dragover', (e) => this.onTaskDragOver(e));
                taskEl.addEventListener('drop', (e) => this.onTaskDrop(e, wp.id));
                taskEl.addEventListener('dragend', (e) => this.onTaskDragEnd(e));
                taskEl.addEventListener('dragleave', (e) => this.onTaskDragLeave(e));
                
                const taskNameEl = document.createElement('span');
                taskNameEl.className = 'task-name';
                taskNameEl.textContent = task.name;
                taskNameEl.title = task.name;
                taskNameEl.style.cursor = 'pointer';
                taskNameEl.onclick = () => this.startEditingTaskName(taskNameEl, task.id, task.name);
                
                // 設定アイコン（3点リーダー）
                const taskMenuBtn = document.createElement('button');
                taskMenuBtn.className = 'task-menu-btn';
                taskMenuBtn.textContent = '⋮';
                taskMenuBtn.title = this.currentLanguage === 'ja' ? '設定' : 'Settings';
                taskMenuBtn.onclick = (e) => {
                    e.stopPropagation();
                    this.toggleTaskMenu(taskEl, task.id, wp.id);
                };
                
                taskEl.appendChild(taskNameEl);
                taskEl.appendChild(taskMenuBtn);
                tasksEl.appendChild(taskEl);
            });
            wpEl.appendChild(tasksEl);
            
            taskList.appendChild(wpEl);
        });

        // 工程追加モード時のプレースホルダーを追加
        if (this.addWorkPackageMode) {
            const placeholder = document.createElement('div');
            placeholder.className = 'workpackage-item add-wp-placeholder';
            placeholder.textContent = this.currentLanguage === 'ja' ? '+ ここに工程を追加' : '+ Add WP here';
            placeholder.addEventListener('click', () => {
                this.addWorkPackageAtIndex(this.workPackages.length);
            });
            taskList.appendChild(placeholder);
        }

        const hint = document.createElement('div');
        hint.style.padding = '10px';
        hint.style.fontSize = '12px';
        hint.style.color = '#999';
        hint.style.textAlign = 'center';
        if (this.currentLanguage === 'ja') {
            hint.textContent = `${this.workPackages.length}/${this.maxWorkPackages} 工程`;
        } else {
            hint.textContent = `${this.workPackages.length}/${this.maxWorkPackages} work packages`;
        }
        taskList.appendChild(hint);
    }

    renderGantt() {
        const ganttContent = document.getElementById('ganttContent');
        ganttContent.innerHTML = '';

        const year = this.currentDate.getFullYear();
        const month = this.currentDate.getMonth();
        const startDate = new Date(year, month, 1);

        // DocumentFragmentで一括追加（レイアウト再計算を最小化）
        const fragment = document.createDocumentFragment();

        if (this.showTasks) {
            // タスク表示オン：各工程ごとにヘッダー行（バーなし）を作成、その後、各タスク行を作成
            this.workPackages.forEach((wp, wpIndex) => {
                // 工程ヘッダー行（バーなし）
                const headerRowEl = this.createGanttRow(null, startDate, `wp-${wp.id}`);
                fragment.appendChild(headerRowEl);

                // 各タスク行（バー表示）
                wp.tasks.forEach((task, taskIndex) => {
                    const rowEl = this.createGanttRow(task, startDate, `task-${task.id}`);
                    fragment.appendChild(rowEl);
                });
            });
        } else {
            // タスク表示オフ：各工程ごとに1行を作成（従来の方式）
            const unitDays = this.getUnitDays();
            const unitCount = Math.ceil(this.daysInView / unitDays);
            const totalWidth = unitCount * this.pixelPerDay * unitDays;
            const cellWidth = this.pixelPerDay * unitDays;

            this.workPackages.forEach((wp, wpIndex) => {
                const rowEl = document.createElement('div');
                rowEl.className = 'gantt-row';
                rowEl.id = `row-${wp.id}`;
                rowEl.style.width = totalWidth + 'px';

                // 背景グリッドをCSSグラデーションで描画
                const rowBgEl = document.createElement('div');
                rowBgEl.className = 'gantt-row-bg';
                rowBgEl.style.width = totalWidth + 'px';
                rowBgEl.style.backgroundImage = `repeating-linear-gradient(to right, transparent, transparent ${cellWidth - 1}px, #ecf0f1 ${cellWidth - 1}px, #ecf0f1 ${cellWidth}px)`;
                rowBgEl.style.backgroundSize = `${cellWidth}px 100%`;
                rowEl.appendChild(rowBgEl);

                const barContainer = document.createElement('div');
                barContainer.className = 'gantt-bar-container';
                barContainer.style.width = totalWidth + 'px';

                // この工程の全タスクを描画
                wp.tasks.forEach((task, index) => {
                    if (this.isTaskInView(task, startDate)) {
                        const barEl = this.createBar(task, startDate, index, wp.tasks.length);
                        barContainer.appendChild(barEl);
                    }
                });

                rowEl.appendChild(barContainer);
                fragment.appendChild(rowEl);
            });
        }

        ganttContent.appendChild(fragment);
    }

    createGanttRow(task, startDate, rowId) {
        const unitDays = this.getUnitDays();
        const unitCount = Math.ceil(this.daysInView / unitDays);
        const totalWidth = unitCount * this.pixelPerDay * unitDays;
        const cellWidth = this.pixelPerDay * unitDays;

        const rowEl = document.createElement('div');
        rowEl.className = 'gantt-row';
        rowEl.id = `${rowId}`;
        rowEl.style.width = totalWidth + 'px';

        // 背景グリッドをCSSグラデーションで描画（DOM要素を大量に作らない）
        const rowBgEl = document.createElement('div');
        rowBgEl.className = 'gantt-row-bg';
        rowBgEl.style.width = totalWidth + 'px';
        rowBgEl.style.backgroundImage = `repeating-linear-gradient(to right, transparent, transparent ${cellWidth - 1}px, #ecf0f1 ${cellWidth - 1}px, #ecf0f1 ${cellWidth}px)`;
        rowBgEl.style.backgroundSize = `${cellWidth}px 100%`;
        rowEl.appendChild(rowBgEl);

        const barContainer = document.createElement('div');
        barContainer.className = 'gantt-bar-container';
        barContainer.style.width = totalWidth + 'px';

        // このタスクを描画（taskがnullの場合はバーを表示しない）
        if (task && this.isTaskInView(task, startDate)) {
            const barEl = this.createBar(task, startDate, 0, 1);
            barContainer.appendChild(barEl);
        }

        rowEl.appendChild(barContainer);
        return rowEl;
    }

    createBar(task, viewStartDate, index, totalTasksInWp) {
        const barEl = document.createElement('div');
        barEl.className = 'gantt-bar';
        barEl.dataset.taskId = task.id;
        barEl.draggable = false;

        // バーの位置と幅を計算
        const daysFromStart = Math.ceil((task.startDate - viewStartDate) / (1000 * 60 * 60 * 24));
        const left = Math.max(0, daysFromStart * this.pixelPerDay);
        const width = Math.min(task.duration * this.pixelPerDay, (this.daysInView - Math.max(0, daysFromStart)) * this.pixelPerDay);

        barEl.style.left = left + 'px';
        barEl.style.width = width + 'px';
        barEl.style.top = '50%';
        barEl.style.transform = 'translateY(-50%)';
        barEl.style.zIndex = totalTasksInWp - index;

        const labelEl = document.createElement('span');
        labelEl.className = 'bar-label';
        labelEl.textContent = task.name;
        
        // バーの幅に応じてフォントサイズを調整
        if (width < 40) {
            labelEl.style.fontSize = '9px';
        } else if (width < 80) {
            labelEl.style.fontSize = '10px';
        }
        
        barEl.appendChild(labelEl);

        // リサイズハンドル
        const leftEdge = document.createElement('div');
        leftEdge.className = 'gantt-bar-edge left';
        barEl.appendChild(leftEdge);

        const rightEdge = document.createElement('div');
        rightEdge.className = 'gantt-bar-edge right';
        barEl.appendChild(rightEdge);

        return barEl;
    }

    isTaskInView(task, viewStartDate) {
        const viewEndDate = new Date(viewStartDate);
        viewEndDate.setDate(viewEndDate.getDate() + this.daysInView);

        const taskEndDate = new Date(task.startDate);
        taskEndDate.setDate(taskEndDate.getDate() + task.duration);

        return task.startDate < viewEndDate && taskEndDate > viewStartDate;
    }

    handleAddTaskClick(e) {
        // クリックされた行（工程）を特定
        const ganttRowEl = e.target.closest('.gantt-row');
        if (!ganttRowEl) {

            return;
        }

        const rowId = ganttRowEl.id;
        let workPackageId = null;

        // workpackage-idを抽出
        if (rowId.startsWith('wp-')) {
            workPackageId = parseInt(rowId.replace('wp-', ''));
        } else if (rowId.startsWith('task-')) {
            // タスク行の場合、そのタスクから工程を見つける
            const taskId = parseInt(rowId.replace('task-', ''));
            for (let wp of this.workPackages) {
                if (wp.tasks.find(t => t.id === taskId)) {
                    workPackageId = wp.id;
                    break;
                }
            }
        } else if (rowId.startsWith('row-')) {
            // タスク非表示モードの場合
            workPackageId = parseInt(rowId.replace('row-', ''));
        }

        if (!workPackageId) {

            return;
        }

        const wp = this.workPackages.find(w => w.id === workPackageId);
        if (!wp) {

            return;
        }

        // クリック位置から日付を計算
        const ganttContent = document.getElementById('ganttContent');
        const containerRect = ganttContent.getBoundingClientRect();
        const scrollLeft = ganttContent.scrollLeft;
        const clickX = e.clientX - containerRect.left + scrollLeft;
        
        const viewStartDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1);
        const daysFromStart = Math.floor(clickX / this.pixelPerDay);
        
        const taskStartDate = new Date(viewStartDate);
        taskStartDate.setDate(taskStartDate.getDate() + daysFromStart);



        // 新しいタスクを追加
        const newId = Math.max(0, ...wp.tasks.map(t => t.id)) + 1;
        const newTask = {
            id: newId,
            name: this.t('newTask'),
            startDate: taskStartDate,
            duration: 3 // 3日間
        };
        
        wp.tasks.push(newTask);
        this.saveData();
        
        // タスク追加モードを解除（toggleAddTaskModeのみを呼び出す）
        this.toggleAddTaskMode();
        
        this.render();
    }

    onMouseDown(e) {
        // タスク追加モードの場合
        if (this.addTaskMode) {
            this.handleAddTaskClick(e);
            return;
        }

        const barEl = e.target.closest('.gantt-bar');
        if (!barEl) return;

        const taskId = parseInt(barEl.dataset.taskId);
        let task = null;
        
        for (let wp of this.workPackages) {
            task = wp.tasks.find(t => t.id === taskId);
            if (task) break;
        }
        
        if (!task) return;

        const isLeftEdge = e.target.closest('.gantt-bar-edge.left');
        const isRightEdge = e.target.closest('.gantt-bar-edge.right');

        const viewStartDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1);
        const barRect = barEl.getBoundingClientRect();
        const containerRect = document.getElementById('ganttContent').getBoundingClientRect();

        this.dragData = {
            taskId: taskId,
            startX: e.clientX,
            originalLeft: barRect.left - containerRect.left,
            originalWidth: barRect.width,
            mode: isLeftEdge ? 'resize-left' : isRightEdge ? 'resize-right' : 'move',
            originalStartDate: new Date(task.startDate),
            originalDuration: task.duration,
            viewStartDate: viewStartDate,
            containerLeft: containerRect.left,
            scrollLeft: document.getElementById('ganttContent').scrollLeft
        };

        barEl.classList.add('dragging');
    }

    onMouseMove(e) {
        if (!this.dragData) return;

        let task = null;
        for (let wp of this.workPackages) {
            task = wp.tasks.find(t => t.id === this.dragData.taskId);
            if (task) break;
        }
        if (!task) return;

        const deltaX = e.clientX - this.dragData.startX;
        const daysPerPixel = 1 / this.pixelPerDay;

        if (this.dragData.mode === 'move') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.startDate = new Date(this.dragData.originalStartDate);
            task.startDate.setDate(task.startDate.getDate() + deltaDays);
        } else if (this.dragData.mode === 'resize-left') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.startDate = new Date(this.dragData.originalStartDate);
            task.startDate.setDate(task.startDate.getDate() + deltaDays);
            task.duration = Math.max(1, this.dragData.originalDuration - deltaDays);
        } else if (this.dragData.mode === 'resize-right') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.duration = Math.max(1, this.dragData.originalDuration + deltaDays);
        }

        this.renderGantt();
    }

    onMouseUp(e) {
        if (this.dragData) {
            const barEl = document.querySelector(`.gantt-bar[data-task-id="${this.dragData.taskId}"]`);
            if (barEl) {
                barEl.classList.remove('dragging');
            }
            this.dragData = null;
            this.saveData();
        }
    }

    // ===== タッチイベント（編集モード時のみバーをドラッグ） =====
    onTouchStart(e) {
        if (!this.editMode) return; // 編集モードでなければ何もしない

        const barEl = e.target.closest('.gantt-bar');
        if (!barEl) return;

        e.preventDefault(); // スクロールを防止

        const touch = e.touches[0];
        const taskId = parseInt(barEl.dataset.taskId);
        let task = null;
        for (let wp of this.workPackages) {
            task = wp.tasks.find(t => t.id === taskId);
            if (task) break;
        }
        if (!task) return;

        const isLeftEdge = e.target.closest('.gantt-bar-edge.left');
        const isRightEdge = e.target.closest('.gantt-bar-edge.right');
        const viewStartDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1);
        const barRect = barEl.getBoundingClientRect();
        const containerRect = document.getElementById('ganttContent').getBoundingClientRect();

        this.dragData = {
            taskId: taskId,
            startX: touch.clientX,
            originalLeft: barRect.left - containerRect.left,
            originalWidth: barRect.width,
            mode: isLeftEdge ? 'resize-left' : isRightEdge ? 'resize-right' : 'move',
            originalStartDate: new Date(task.startDate),
            originalDuration: task.duration,
            viewStartDate: viewStartDate,
            containerLeft: containerRect.left,
            scrollLeft: document.getElementById('ganttContent').scrollLeft
        };

        barEl.classList.add('dragging');
    }

    onTouchMove(e) {
        if (!this.editMode || !this.dragData) return;

        e.preventDefault();

        const touch = e.touches[0];
        let task = null;
        for (let wp of this.workPackages) {
            task = wp.tasks.find(t => t.id === this.dragData.taskId);
            if (task) break;
        }
        if (!task) return;

        const deltaX = touch.clientX - this.dragData.startX;
        const daysPerPixel = 1 / this.pixelPerDay;

        if (this.dragData.mode === 'move') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.startDate = new Date(this.dragData.originalStartDate);
            task.startDate.setDate(task.startDate.getDate() + deltaDays);
        } else if (this.dragData.mode === 'resize-left') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.startDate = new Date(this.dragData.originalStartDate);
            task.startDate.setDate(task.startDate.getDate() + deltaDays);
            task.duration = Math.max(1, this.dragData.originalDuration - deltaDays);
        } else if (this.dragData.mode === 'resize-right') {
            const deltaDays = Math.round(deltaX * daysPerPixel);
            task.duration = Math.max(1, this.dragData.originalDuration + deltaDays);
        }

        this.renderGantt();
    }

    onTouchEnd(e) {
        if (!this.editMode) return;

        if (this.dragData) {
            const barEl = document.querySelector(`.gantt-bar[data-task-id="${this.dragData.taskId}"]`);
            if (barEl) {
                barEl.classList.remove('dragging');
            }
            this.dragData = null;
            this.saveData();
        }
    }

    onWorkPackageDragStart(e, wpId) {
        this.draggedWorkPackageId = wpId;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', wpId.toString());
        e.target.classList.add('dragging-wp');
    }

    onWorkPackageDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        
        if (e.target.closest('.workpackage-item') && !e.target.closest('.workpackage-item').classList.contains('dragging-wp')) {
            e.target.closest('.workpackage-item').classList.add('drag-over-wp');
        }
    }

    onWorkPackageDragLeave(e) {
        if (e.target.closest('.workpackage-item')) {
            e.target.closest('.workpackage-item').classList.remove('drag-over-wp');
        }
    }

    onWorkPackageDrop(e, targetWpId) {
        e.preventDefault();
        e.stopPropagation();
        
        const draggedWpId = this.draggedWorkPackageId;
        
        if (draggedWpId !== targetWpId) {
            const draggedIndex = this.workPackages.findIndex(wp => wp.id === draggedWpId);
            const targetIndex = this.workPackages.findIndex(wp => wp.id === targetWpId);
            
            if (draggedIndex !== -1 && targetIndex !== -1) {
                // 工程の順序を入れ替え
                const temp = this.workPackages[draggedIndex];
                this.workPackages[draggedIndex] = this.workPackages[targetIndex];
                this.workPackages[targetIndex] = temp;
                
                this.saveData();
                this.render();
            }
        }
        
        // ドラッグオーバースタイルを削除
        const allWpItems = document.querySelectorAll('.workpackage-item');
        allWpItems.forEach(item => {
            item.classList.remove('drag-over-wp');
        });
    }

    onWorkPackageDragEnd(e) {
        this.draggedWorkPackageId = null;
        
        // すべてのドラッグ中スタイルを削除
        const allWpItems = document.querySelectorAll('.workpackage-item');
        allWpItems.forEach(item => {
            item.classList.remove('dragging-wp');
            item.classList.remove('drag-over-wp');
        });
    }

    // タスクドラッグ&ドロップメソッド
    onTaskDragStart(e, taskId, wpId) {

        this.draggedTaskId = taskId;
        this.draggedTaskSourceWpId = wpId;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', `task-${taskId}`);
        e.target.closest('.task-item').classList.add('dragging-task');
    }

    onTaskDragOver(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        
        const taskItem = e.target.closest('.task-item');
        const wpItem = e.target.closest('.workpackage-item');
        
        // タスク上での場合
        if (taskItem && !taskItem.classList.contains('dragging-task')) {
            taskItem.classList.add('drag-over-task');
        }
        // 工程上での場合
        else if (wpItem && !wpItem.classList.contains('dragging-task')) {
            wpItem.classList.add('drag-over-wp');
        }
    }

    onTaskDragLeave(e) {
        const taskItem = e.target.closest('.task-item');
        const wpItem = e.target.closest('.workpackage-item');
        
        if (taskItem) {
            taskItem.classList.remove('drag-over-task');
        }
        if (wpItem) {
            wpItem.classList.remove('drag-over-wp');
        }
    }

    onTaskDrop(e, targetWpId) {
        e.preventDefault();
        e.stopPropagation();
        
        const draggedTaskId = this.draggedTaskId;
        const sourceWpId = this.draggedTaskSourceWpId;
        
        if (!draggedTaskId || !sourceWpId) {
            this.clearDragOverStyles();
            return;
        }
        

        
        // ドラッグ元とドラッグ先の工程を取得
        const sourceWp = this.workPackages.find(wp => wp.id === sourceWpId);
        const targetWp = this.workPackages.find(wp => wp.id === targetWpId);
        
        if (!sourceWp || !targetWp) {
            // ドラッグオーバースタイルを削除
            this.clearDragOverStyles();
            return;
        }
        
        // ドラッグ元のタスクを見つける
        const taskIndex = sourceWp.tasks.findIndex(t => t.id === draggedTaskId);
        if (taskIndex === -1) {
            // ドラッグオーバースタイルを削除
            this.clearDragOverStyles();
            return;
        }
        
        // 同じ工程への移動の場合、タスクの順序を変更
        if (sourceWpId === targetWpId) {
            // ドロップ対象のタスクを見つける
            const targetTaskEl = e.target.closest('.task-item');
            if (targetTaskEl && targetTaskEl.dataset.taskId) {
                const targetTaskId = parseInt(targetTaskEl.dataset.taskId);
                if (!isNaN(targetTaskId) && draggedTaskId !== targetTaskId) {
                    let targetTaskIndex = sourceWp.tasks.findIndex(t => t.id === targetTaskId);
                    
                    if (targetTaskIndex !== -1) {
                        // タスクを移動
                        const task = sourceWp.tasks.splice(taskIndex, 1)[0];
                        
                        // インデックス調整：削除後、ターゲットのインデックスが変わる場合がある
                        if (taskIndex < targetTaskIndex) {
                            targetTaskIndex--;
                        }
                        
                        sourceWp.tasks.splice(targetTaskIndex, 0, task);
                        
                        this.saveData();
                        this.render();
                        return;
                    }
                }
            }
            // ドラッグオーバースタイルを削除
            this.clearDragOverStyles();
            return;
        }
        
        // 異なる工程への移動の場合
        // タスクを元の工程から削除
        const task = sourceWp.tasks.splice(taskIndex, 1)[0];
        
        // タスクをターゲット工程に追加
        targetWp.tasks.push(task);
        
        this.saveData();
        this.render();
    }

    onTaskDragEnd(e) {
        this.draggedTaskId = null;
        this.draggedTaskSourceWpId = null;
        
        // すべてのドラッグ中スタイルを削除
        this.clearDragOverStyles();
    }

    clearDragOverStyles() {
        const allTaskItems = document.querySelectorAll('.task-item');
        allTaskItems.forEach(item => {
            item.classList.remove('dragging-task');
            item.classList.remove('drag-over-task');
        });
        
        const allWpItems = document.querySelectorAll('.workpackage-item');
        allWpItems.forEach(item => {
            item.classList.remove('drag-over-wp');
        });
    }

    async exportToExcel() {
        // ExcelJSライブラリが読み込まれているかチェック
        if (typeof ExcelJS === 'undefined') {
            alert('エクセルエクスポート機能の読み込みに失敗しました。ページをリロードしてください。');
            return;
        }

        // ワークブックを作成
        const workbook = new ExcelJS.Workbook();
        const sheetName = this.currentLanguage === 'ja' ? 'ガントチャート' : 'Gantt Chart';
        const worksheet = workbook.addWorksheet(sheetName);

        // 表示期間を計算（現在表示中の月から12ヶ月分）
        const startDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth(), 1);
        const endDate = new Date(startDate);
        endDate.setMonth(endDate.getMonth() + this.monthsInView);
        
        // 日付配列を作成
        const dates = [];
        const currentDay = new Date(startDate);
        while (currentDay < endDate) {
            dates.push(new Date(currentDay));
            currentDay.setDate(currentDay.getDate() + 1);
        }

        // 列幅を設定
        worksheet.getColumn(1).width = 15; // 工程名
        worksheet.getColumn(2).width = 25; // タスク名
        for (let i = 0; i < dates.length; i++) {
            worksheet.getColumn(i + 3).width = 3; // 日付列
        }

        // ヘッダー行1: 年月
        const monthRow = worksheet.getRow(1);
        monthRow.getCell(1).value = '';
        monthRow.getCell(2).value = '';
        let lastMonth = '';
        let mergeStart = 3;
        dates.forEach((date, index) => {
            const monthStr = `${date.getFullYear()}/${date.getMonth() + 1}`;
            const col = index + 3;
            
            if (monthStr !== lastMonth) {
                if (lastMonth && mergeStart < col) {
                    worksheet.mergeCells(1, mergeStart, 1, col - 1);
                }
                monthRow.getCell(col).value = monthStr;
                mergeStart = col;
                lastMonth = monthStr;
            }
        });
        if (mergeStart < dates.length + 2) {
            worksheet.mergeCells(1, mergeStart, 1, dates.length + 2);
        }
        
        // ヘッダー行1のスタイル
        monthRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFCCCCCC' }
            };
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // ヘッダー行2: 日
        const dayRow = worksheet.getRow(2);
        dayRow.getCell(1).value = this.currentLanguage === 'ja' ? '工程' : 'WP';
        dayRow.getCell(2).value = this.currentLanguage === 'ja' ? 'タスク' : 'Task';
        dates.forEach((date, index) => {
            dayRow.getCell(index + 3).value = date.getDate();
        });
        
        // ヘッダー行2のスタイル
        dayRow.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFEEEEEE' }
            };
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });

        // タスクデータ行
        let rowIndex = 3;
        const taskColors = [
            'FF4472C4', 'FFED7D31', 'FFA5A5A5', 'FFFFC000', 
            'FF5B9BD5', 'FF70AD47', 'FF9E480E', 'FF636363'
        ];
        
        this.workPackages.forEach((wp, wpIndex) => {
            const color = taskColors[wpIndex % taskColors.length];
            
            wp.tasks.forEach((task, taskIndex) => {
                const row = worksheet.getRow(rowIndex);
                
                // 工程名とタスク名
                row.getCell(1).value = taskIndex === 0 ? wp.name : '';
                row.getCell(2).value = task.name;
                
                // タスク期間を計算
                const taskStart = new Date(task.startDate);
                const taskEnd = new Date(task.startDate);
                taskEnd.setDate(taskEnd.getDate() + task.duration - 1);
                
                // 各日付セルを処理
                dates.forEach((date, dateIndex) => {
                    const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
                    const startOnly = new Date(taskStart.getFullYear(), taskStart.getMonth(), taskStart.getDate());
                    const endOnly = new Date(taskEnd.getFullYear(), taskEnd.getMonth(), taskEnd.getDate());
                    
                    const cell = row.getCell(dateIndex + 3);
                    cell.value = ''; // セルは空にする
                    
                    if (dateOnly >= startOnly && dateOnly <= endOnly) {
                        // タスク期間内のセルに背景色を設定
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: color }
                        };
                    }
                    
                    // 罫線を設定
                    cell.border = {
                        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                        bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
                        right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
                    };
                });
                
                // 工程名とタスク名のセルスタイル
                row.getCell(1).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                row.getCell(2).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                
                rowIndex++;
            });
        });

        // 行の高さを設定
        worksheet.eachRow((row) => {
            row.height = 20;
        });

        // ファイルを出力
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const fileName = `gantt_chart_${new Date().toISOString().split('T')[0]}.xlsx`;
        saveAs(blob, fileName);
    }
}

// グローバルインスタンス
let gantt;

// ページ読み込み時に初期化
document.addEventListener('DOMContentLoaded', () => {
    gantt = new GanttChart();
});
