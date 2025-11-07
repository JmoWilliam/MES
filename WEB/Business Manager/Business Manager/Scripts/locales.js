const toDoListLocales = {
    // calendar
    calendar: {
        monthFull: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
        monthShort: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
        dayFull: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"],
        dayShort: ["日", "一", "二", "三", "四", "五", "六"],

        clear: "清除",
        done: "完成",
        today: "今天",
    },
    // To Do List
    todo: {
        // Toolbar
        "No project": "沒有項目",
        "Search project": "搜尋項目",
        "Add project": "新增項目",
        "Rename project": "重新命名項目",
        "Delete project": "刪除項目",

        // Task
        "Add task below": "下方新增任務",
        "Add subtask": "新增子任務",
        "Set due date": "設定截止日期",
        "Indent": "縮排",
        "Unindent": "移除縮排",
        "Assign to": "指派",
        "Move to": "移動",
        "Duplicate": "複製副本",
        "Copy": "複製",
        "Paste": "貼上",
        "Delete": "刪除",

        // Shortcut
        "Enter": "Enter",
        "Tab": "Tab",
        "Shift+Tab": "Shift+Tab",
        "Ctrl+D": "Ctrl+D",
        "Ctrl+C": "Ctrl+C",
        "Ctrl+V": "Ctrl+V",
        // For Mac OS
        "CMD+D": "CMD+D",
        "CMD+C": "CMD+C",
        "CMD+V": "CMD+V",

        // Editor
        "Type you want": "等待輸入",

        // Other
        "Search": "搜尋",
        "Add task": "新增任務",
        "New project": "新增項目",
    }
};

const ganttLocales = {
    date: {
        month_full: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
        month_short: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
        day_full: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"],
        day_short: ["日", "一", "二", "三", "四", "五", "六"]
    },
    labels: {
        new_task: "新任務",
        icon_save: "儲存",
        icon_cancel: "取消",
        icon_details: "詳細資料",
        icon_edit: "修改",
        icon_delete: "刪除",
        gantt_save_btn: "New Label",
        gantt_cancel_btn: "New Label",
        gantt_delete_btn: "New Label",
        confirm_closing: "任務尚未儲存，是否關閉?",
        confirm_deleting: "是否刪除任務?",
        section_description: "任務名稱",
        section_time: "時間範圍",
        section_type: "類型",

        /* grid columns */
        column_wbs: "WBS",
        column_text: "任務名稱",
        column_start_date: "開始時間",
        column_end_date: "結束時間",
        column_duration: "持續時間",
        column_add: "",

        /* link confirmation */
        link: "", //連結
        confirm_link_deleting: "", //是否要刪除?
        link_start: " (start)",
        link_end: " (end)",

        type_task: "任務",
        type_project: "專案",
        type_milestone: "里程碑",

        minutes: "Minutes",
        hours: "Hours",
        days: "Days",
        weeks: "Week",
        months: "Months",
        years: "Years",

        /* message popup */
        message_ok: "確認",
        message_cancel: "取消",

        /* constraints */
        section_constraint: "Constraint",
        constraint_type: "Constraint type",
        constraint_date: "Constraint date",
        asap: "As Soon As Possible",
        alap: "As Late As Possible",
        snet: "Start No Earlier Than",
        snlt: "Start No Later Than",
        fnet: "Finish No Earlier Than",
        fnlt: "Finish No Later Than",
        mso: "Must Start On",
        mfo: "Must Finish On",

        /* resource control */
        resources_filter_placeholder: "type to filter",
        resources_filter_label: "hide empty"
    }
};

const richTextLocales = {
    apply: "套用",
    undo: "復原",
    redo: "取消復原",
    selectFontFamily: "字型",
    selectFontSize: "字型大小",
    selectFormat: "樣式",
    selectTextColor: "字型色彩",
    selectTextBackground: "背景色彩",
    markBold: "粗體",
    markItalic: "斜體",
    markStrike: "刪除線",
    markUnderline: "底線",
    alignLeft: "靠左對齊",
    alignCenter: "置中",
    alignRight: "靠右對齊",
    addLink: "超連結",
    clearFormat: "清除格式",
    fullscreen: "全螢幕",
    removeLink: "移除超連結",
    edit: "編輯",
    h1: "標題 1",
    h2: "標題 2",
    h3: "標題 3",
    h4: "標題 4",
    h5: "標題 5",
    h6: "標題 6",
    p: "段落",
    blockquote: "引用",
    stats: "Statistics",
    chars: "chars",
    charsExlSpace: "charsExlSpace",
    words: "words",
    text: "Text",
    link: "Link"
};

const calendarLocales = {
    monthsShort: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
    months: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
    daysShort: ["日", "一", "二", "三", "四", "五", "六"],
    days: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"],
    cancel: "取消"
};

const eventCalendarLocales = {
    /*translations and settings of the calendar*/
    calendar: {
        monthFull: ["一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月"],
        monthShort: ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"],
        dayFull: ["星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"],
        dayShort: ["日", "一", "二", "三", "四", "五", "六"],
        hours: "時",
        minutes: "分",
        done: "完成",
        clear: "清除",
        today: "今天",
        am: ["am", "AM"],
        pm: ["pm", "PM"],
        weekStart: 1, // defines a first day of week (Sunday by default)
    },
    /*translations of the Event Calendar labels*/
    scheduler: {
        "New Event": "新事件",
        "Add description": "新增描述",
        "Create event": "新增事件",
        "Edit event": "編輯事件",
        "Delete event": "刪除事件",
        "Event name": "事件名稱",
        "Start date": "開始日期",
        "End date": "結束日期",
        "All day": "整天",
        "No events": "沒有事件",
        Type: "類型",
        Description: "描述",
        Today: "今天",
        Day: "日",
        Week: "週",
        Month: "月",
        Timeline: "時間線",
        Calendars: "Calendars",
        hourFormat: "H",
        minuteFormat: "mm",
        monthFormat: "EEE",
        dateFormat: "EEE, d",
        agendaDateFormat: "MMMM d EEEE",
        unassignFormat: "d MMM yyyy",
        Color: "顏色",
        Delete: "刪除",
        Edit: "編輯",
        Calendar: "Calendar",
        New: "新",
        Name: "名稱",
        Save: "儲存",
        Add: "新增",
        Event: "事件",
        confirmDelete: "Are you sure you want to delete {item}?",
        confirmUnsaved: "You have unsaved changes!  Do you want to discard them?",
        "Repeat event": "Repeat event",
        viewAll: "+{count} more",
        Never: "從不",
        Every: "每個",
        Workdays: "On work days",
        Year: "年",
        Custom: "Custom",
        Ends: "Ends",
        After: "After",
        "On date": "On date",
        events: "events",
        "recurring event": "recurring event",
        all: "每個事件",
        future: "This and following events",
        only: "This event",
        recurringActionError: "Start date cannot be after recurrent expiry date",
        Assignees: "Assignees",
        "Recurring events": "Recurring events",
        "Single events": "Single events",
        recurringEveryMonthDay: "每個 {date}",
        recurringEveryMonthPos: "每個 {pos} {weekDay}",
        recurringEveryYearDay: "每個 {month}/{date}",
        recurringEveryYearPos: "每個 {pos} {weekDay} of {month}"
    },
    /*translations of the WX core elements*/
    core: {
        ok: "確定",
        cancel: "取消"
    }
};

const timepickerLocales = {
    hours: "時",
    minutes: "分",
    save: "確定"
};

const diagramLocales = {
    applyAll: "套用",
    exportData: "匯出",
    importData: "匯入",
    resetChanges: "重置",
    autoLayout: "自動排版",
    orthogonal: "Orthogonal", // added in v5.0
    radial: "Radial", // added in v5.0

    shapeSections: "Shapes",
    groupSections: "Groups",
    swimlaneSections: "Swimlanes",

    gridStep: "Grid step",
    arrange: "Arrange",
    position: "Position",
    size: "Size",
    color: "Color",
    title: "Title",
    text: "Text",
    image: "Image",
    fill: "Fill",
    textProps: "Text",
    stroke: "Stroke",

    headerText: "Header text",
    headerFill: "Header fill",
    headerStyle: "Header style",
    headerPosition: "Header position",
    headerEnable: "Header",
    subHeaderRowsEnable: "Subheader rows",
    subHeaderColsEnable: "Subheader cols",

    positionOptions: {
        top: "top",
        bottom: "bottom",
        left: "left",
        right: "right",
    },
    switchOptions: {
        on: "on",
        off: "off",
    },

    menuDeleteRow: "Delete row",
    menuDeleteCol: "Delete column",
    menuMoveColumnRight: "Move column right",
    menuMoveColumnLeft: "Move column left",
    menuMoveRowUp: "Move row up",
    menuMoveRowDown: "Move row down",
    menuAddRowUp: "Add row up",
    menuAddRowDown: "Add row down",
    menuAddColumnRight: "Add column right",
    menuAddColumnLeft: "Add column left",
    menuDelete: "Delete",
    menuAddPartner: "Add partner",
    menuAddAssistant: "Add assistant",
    menuAlignChildrenVertically: "Align children vertically",
    menuAlignChildrenHorizontally: "Align children horizontally",

    imageUpload: "Click to upload",
    emptyState: "Select a shape or a connector",

     // the following locale options were added in v5.0
    alignHorizontalLeft: "Align left",
    alignHorizontalCenter: "Align horizontal centers",
    alignHorizontalRight: "Align right",
    alignHorizontalDistribution: "Distribute horizontal spacing",
    alignVerticalDistribution: "Distribute vertical spacing",
    alignVerticalTop: "Align top",
    alignVerticalMiddle: "Align vertical centers",
    alignVerticalBottom: "Align bottom",

    addShape: "Add shape",
    menu: "Menu",
    remove: "移除",
    addLeftShape: "Add left shape",
    addRightShape: "Add right shape",

    lineTextAutoPositionEnable: "Enable text autoposition",
    lineTextAutoPositionDisable: "Disable text autoposition",
    addLineText: "Add text",
    addColumnLast: "Add column",
    addRowLast: "Add row",
    copy: "複製",
    connect: "串聯",
    removePoint: "Delete point",

};