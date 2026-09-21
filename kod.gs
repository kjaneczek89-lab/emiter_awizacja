const SPREADSHEET_ID = '1MplrlnC81gNaGKUVm0k_z5ay82ytJks6mUBDu-jSmhw';

// 1. Zwykłe wejście z przeglądarki na link skryptu - by sprawdzić, czy działa
function doGet() {
  return ContentService.createTextOutput("System Awizacji - API działa poprawnie.");
}

// 2. Odbieranie zapytań z zewnętrznej strony (Vercel)
function doPost(e) {
  try {
    // Odczytujemy dane przesłane z Vercela
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result = {};

    // Kierujemy do odpowiedniej funkcji w zależności od 'action'
    if (action === 'verifyLogin') {
      result = verifyLogin(request.login, request.haslo);
    } else if (action === 'createAccount') {
      result = createAccount(request.login, request.haslo, request.imieNazwisko, request.firma, request.rola);
    } else if (action === 'changePassword') {
      result = changePassword(request.userId, request.newPassword);
    } else if (action === 'zapiszAwizacje') {
      result = zapiszDaneAwizacji(request.awizacja, request.towary);
    } else if (action === 'pobierzAwizacje') {
      result = pobierzAwizacje(request.rola, request.userId, request.userFirma);
    } else {
      result = { success: false, message: "Nieznana akcja API." };
    }

    // Zwracamy wynik w formacie JSON
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "Błąd serwera: " + error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 3. Funkcja szyfrująca hasła (SHA-256)
function hashPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

// 4. Tworzenie konta
function createAccount(login, plainPassword, imieNazwisko, firma, rola) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === login) {
      return { success: false, message: "Użytkownik o tym loginie już istnieje!" };
    }
  }
  
  const newId = "U-" + (new Date().getTime());
  const hashedPassword = hashPassword(plainPassword);
  sheet.appendRow([newId, login, hashedPassword, imieNazwisko, firma, rola, true]);
  
  return { success: true, message: "Konto zostało utworzone pomyślnie." };
}

// 5. Weryfikacja logowania
function verifyLogin(login, plainPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  const hashedPassword = hashPassword(plainPassword);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === login && data[i][2] === hashedPassword) {
      let requiresChange = (data[i][6] === true || data[i][6] === 'true' || data[i][6] === 'TRUE');
      return {
        success: true,
        user: { id: data[i][0], login: data[i][1], imieNazwisko: data[i][3], firma: data[i][4], rola: data[i][5], wymagaZmiany: requiresChange }
      };
    }
  }
  return { success: false, message: "Nieprawidłowy login lub hasło." };
}

// 6. Zmiana hasła
function changePassword(userId, newPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      const newHash = hashPassword(newPassword);
      sheet.getRange(i + 1, 3).setValue(newHash);
      sheet.getRange(i + 1, 7).setValue(false); 
      return { success: true };
    }
  }
  return { success: false, message: "Błąd: Nie znaleziono użytkownika w bazie." };
}

// 7. Zapisywanie awizacji
function zapiszDaneAwizacji(awizacja, towary) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const awizacjeSheet = ss.getSheetByName('Awizacje');
  const towarySheet = ss.getSheetByName('Towary');
  
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const idAwizacji = "AW-" + dateStr;
  
  awizacjeSheet.appendRow([
    idAwizacji, awizacja.userId, awizacja.firmaZglaszajaca, awizacja.dataPrzyjazdu,
    awizacja.rejestracja, awizacja.przewoznik, "Planowany", awizacja.kierunek, awizacja.kontrahent 
  ]);
  
  towary.forEach((towar, index) => {
    const idTowaru = idAwizacji + "-T" + (index + 1);
    towarySheet.appendRow([ idTowaru, idAwizacji, towar.nazwa, towar.ilosc, towar.jednostka ]);
  });
  
  return { success: true, message: "Pomyślnie zarejestrowano awizację!" };
}

// 8. Pobieranie Harmonogramu - USUNIĘTO OGRANICZENIA WIDOCZNOŚCI
function pobierzAwizacje(rola, userId, userFirma) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const awizacjeSheet = ss.getSheetByName('Awizacje');
  const towarySheet = ss.getSheetByName('Towary');
  
  const awizDataRange = awizacjeSheet.getDataRange().getValues();
  const towaryDataRange = towarySheet.getDataRange().getValues();
  
  if (awizDataRange.length <= 1) return []; 

  const awizacje = [];
  const towaryMap = {};
  
  for (let i = 1; i < towaryDataRange.length; i++) {
     let idAwiz = towaryDataRange[i][1];
     if (!towaryMap[idAwiz]) towaryMap[idAwiz] = [];
     towaryMap[idAwiz].push({ nazwa: towaryDataRange[i][2], ilosc: towaryDataRange[i][3], jednostka: towaryDataRange[i][4] });
  }

  for (let i = 1; i < awizDataRange.length; i++) {
    let row = awizDataRange[i];
    let idAwizacji = row[0];
    let kontrahent = row[8] ? String(row[8]) : String(row[2]); 
    
    // Brak warunków if(rola==='Admin') - każdy widzi wszystko
    let rawDate = row[3];
    let dateStr = "";
    if (rawDate instanceof Date) {
       dateStr = rawDate.toISOString(); 
    } else {
       dateStr = String(rawDate);
    }
    
    awizacje.push({
      id: idAwizacji,
      firma: row[2], 
      kontrahent: kontrahent,
      kierunek: row[7] || "Dostawa do Emitera",
      data: dateStr, 
      rejestracja: row[4],
      przewoznik: row[5],
      status: row[6],
      towary: towaryMap[idAwizacji] || []
    });
  }
  
  awizacje.sort(function(a, b) { return new Date(a.data) - new Date(b.data); });
  return awizacje;
}
