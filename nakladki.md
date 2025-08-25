function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Анализ расписания')
    .addItem('Найти накладки', 'findScheduleConflicts')
    .addToUi();
}

function findScheduleConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); // Используем активный лист
  const data = sheet.getDataRange().getValues();
  
  // Создаем или очищаем лист для результатов
  let conflictSheet = ss.getSheetByName('Накладки');
  if (!conflictSheet) {
    conflictSheet = ss.insertSheet('Накладки');
  } else {
    conflictSheet.clear();
  }
  
  // Заголовки для листа с накладками
  conflictSheet.getRange(1, 1, 1, 7).setValues([[
    'Тип накладки', 'День', 'Урок', 'Группы', 'Преподаватели', 'Кабинеты', 'Описание'
  ]]);
  
  const headers = data[0]; // Названия групп
  const conflicts = [];
  
  // Объект для хранения данных по времени (день + урок)
  const timeSlots = {};
  
  // Парсим данные из таблицы
  for (let row = 1; row < data.length; row++) {
    const day = data[row][0]; // Колонка A - день
    const lesson = data[row][1]; // Колонка B - номер урока
    
    if (!day || !lesson) continue;
    
    const timeKey = `${day}_${lesson}`;
    if (!timeSlots[timeKey]) {
      timeSlots[timeKey] = [];
    }
    
    // Проходим по колонкам с предметами (начиная с колонки C)
    for (let col = 2; col < data[row].length; col += 2) {
      const subjectCell = data[row][col];
      const roomCell = col + 1 < data[row].length ? data[row][col + 1] : '';
      
      if (!subjectCell || subjectCell.toString().trim() === '') continue;
      
      const groupName = headers[col];
      if (!groupName) continue;
      
      // Извлекаем преподавателя из ячейки предмета
      let teacher = 'Не указан';
      let subject = subjectCell.toString().trim();
      
      if (subject.includes('/')) {
        const parts = subject.split('/');
        subject = parts[0].trim();
        teacher = parts[1].trim();
      }
      
      const room = roomCell ? roomCell.toString().trim() : 'Не указан';
      
      timeSlots[timeKey].push({
        day: day,
        lesson: lesson,
        group: groupName,
        subject: subject,
        teacher: teacher,
        room: room
      });
    }
  }
  
  // Анализируем накладки
  for (const timeKey in timeSlots) {
    const entries = timeSlots[timeKey];
    if (entries.length < 2) continue;
    
    // Проверяем накладки преподавателей (один преподаватель в разных кабинетах)
    const teacherRooms = {};
    for (const entry of entries) {
      if (!teacherRooms[entry.teacher]) {
        teacherRooms[entry.teacher] = new Set();
      }
      teacherRooms[entry.teacher].add(entry.room);
    }
    
    for (const [teacher, rooms] of Object.entries(teacherRooms)) {
      if (rooms.size > 1) {
        const teacherEntries = entries.filter(e => e.teacher === teacher);
        const groups = teacherEntries.map(e => e.group).join(', ');
        const roomsList = Array.from(rooms).join(', ');
        
        conflicts.push([
          'Конфликт преподавателя',
          teacherEntries[0].day,
          teacherEntries[0].lesson,
          groups,
          teacher,
          roomsList,
          `Преподаватель ${teacher} одновременно в ${rooms.size} кабинетах`
        ]);
      }
    }
    
    // Проверяем накладки кабинетов (более 2 групп в одном кабинете)
    const roomGroups = {};
    const roomTeachers = {};
    
    for (const entry of entries) {
      if (!roomGroups[entry.room]) {
        roomGroups[entry.room] = new Set();
        roomTeachers[entry.room] = new Set();
      }
      roomGroups[entry.room].add(entry.group);
      roomTeachers[entry.room].add(entry.teacher);
    }
    
    // Более 2 групп в одном кабинете
    for (const [room, groups] of Object.entries(roomGroups)) {
      if (groups.size > 2) {
        const groupsList = Array.from(groups).join(', ');
        const teachersList = Array.from(roomTeachers[room]).join(', ');
        
        conflicts.push([
          'Перегрузка кабинета (группы)',
          entries[0].day,
          entries[0].lesson,
          groupsList,
          teachersList,
          room,
          `В кабинете ${room} одновременно ${groups.size} групп`
        ]);
      }
    }
    
    // Более 2 преподавателей в одном кабинете
    for (const [room, teachers] of Object.entries(roomTeachers)) {
      if (teachers.size > 2) {
        const teachersList = Array.from(teachers).join(', ');
        const groupsList = Array.from(roomGroups[room]).join(', ');
        
        conflicts.push([
          'Конфликт преподавателей',
          entries[0].day,
          entries[0].lesson,
          groupsList,
          teachersList,
          room,
          `В кабинете ${room} одновременно ${teachers.size} преподавателей`
        ]);
      }
    }
  }
  
  // Записываем результаты
  if (conflicts.length > 0) {
    conflictSheet.getRange(2, 1, conflicts.length, 7).setValues(conflicts);
  } else {
    conflictSheet.getRange(2, 1).setValue('Накладки не найдены');
  }
  
  // Форматируем заголовки
  conflictSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  conflictSheet.autoResizeColumns(1, 7);
  
  SpreadsheetApp.getUi().alert(`Анализ завершен. Найдено ${conflicts.length} накладок.`);
}
