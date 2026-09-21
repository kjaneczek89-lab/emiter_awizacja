const SPREADSHEET_ID = '1MplrlnC81gNaGKUVm0k_z5ay82ytJks6mUBDu-jSmhw';

// 1. Serwowanie pliku HTML (Frontend)
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('System Awizacji - Emiter')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Funkcja szyfrująca hasła (SHA-256)
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

// 3. Tworzenie konta (używane przez Panel Administratora)
function createAccount(login, plainPassword, imieNazwisko, firma, rola) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  // Sprawdzenie, czy login jest już zajęty
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === login) {
      return { success: false, message: "Użytkownik o tym loginie już istnieje!" };
    }
  }
  
  const newId = "U-" + (new Date().getTime());
  const hashedPassword = hashPassword(plainPassword);
  
  // Dodajemy wiersz - kolumna G (true) oznacza wymóg zmiany hasła przy pierwszym logowaniu
  sheet.appendRow([newId, login, hashedPassword, imieNazwisko, firma, rola, true]);
  
  return { success: true, message: "Konto zostało utworzone pomyślnie." };
}

// 4. Weryfikacja logowania
function verifyLogin(login, plainPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  const hashedPassword = hashPassword(plainPassword);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === login && data[i][2] === hashedPassword) {
      
      // Sprawdzamy czy użytkownik musi zmienić hasło (Kolumna G)
      let requiresChange = (data[i][6] === true || data[i][6] === 'true' || data[i][6] === 'TRUE');
      
      return {
        success: true,
        user: {
          id: data[i][0],
          login: data[i][1],
          imieNazwisko: data[i][3],
          firma: data[i][4],
          rola: data[i][5],
          wymagaZmiany: requiresChange
        }
      };
    }
  }
  return { success: false, message: "Nieprawidłowy login lub hasło." };
}

// 5. Zmiana hasła przez użytkownika
function changePassword(userId, newPassword) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Uzytkownicy');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      const newHash = hashPassword(newPassword);
      sheet.getRange(i + 1, 3).setValue(newHash); // Zapis nowego zaszyfrowanego hasła w kolumnie C
      sheet.getRange(i + 1, 7).setValue(false);   // Wyłączenie wymogu zmiany hasła w kolumnie G
      return { success: true };
    }
  }
  return { success: false, message: "Błąd: Nie znaleziono użytkownika w bazie." };
}

// 6. Zapisywanie nowej awizacji i wielu towarów
function zapiszDaneAwizacji(awizacja, towary) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const awizacjeSheet = ss.getSheetByName('Awizacje');
  const towarySheet = ss.getSheetByName('Towary');
  
  // Generowanie unikalnego ID Awizacji
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  const idAwizacji = "AW-" + dateStr;
  
  // Zapis do tabeli 'Awizacje' (W tym kierunek i kontrahent)
  awizacjeSheet.appendRow([
    idAwizacji,
    awizacja.userId,
    awizacja.firmaZglaszajaca, 
    awizacja.dataPrzyjazdu,
    awizacja.rejestracja,
    awizacja.przewoznik,
    "Planowany",         // Kolumna G: Status
    awizacja.kierunek,   // Kolumna H: Dostawa lub Wysyłka
    awizacja.kontrahent  // Kolumna I: Odbiorca / Nadawca (np. Kelvion)
  ]);
  
  // Zapis powiązanych towarów do tabeli 'Towary'
  towary.forEach((towar, index) => {
    const idTowaru = idAwizacji + "-T" + (index + 1);
    towarySheet.appendRow([
      idTowaru,
      idAwizacji,
      towar.nazwa,
      towar.ilosc,
      towar.jednostka
    ]);
  });
  
  return { success: true, message: "Pomyślnie zarejestrowano awizację!" };
}

// 7. Pobieranie Harmonogramu dla widoku listy
function pobierzAwizacje(rola, userId, userFirma) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const awizacjeSheet = ss.getSheetByName('Awizacje');
  const towarySheet = ss.getSheetByName('Towary');
  
  const awizDataRange = awizacjeSheet.getDataRange().getValues();
  const towaryDataRange = towarySheet.getDataRange().getValues();
  
  // Jeśli są same nagłówki (puste arkusze), od razu zwróć pustą listę
  if (awizDataRange.length <= 1) return []; 

  const awizacje = [];
  const towaryMap = {};
  
  // Grupowanie towarów przypisanych do danej awizacji
  for (let i = 1; i < towaryDataRange.length; i++) {
     let idAwiz = towaryDataRange[i][1];
     if (!towaryMap[idAwiz]) towaryMap[idAwiz] = [];
     towaryMap[idAwiz].push({
       nazwa: towaryDataRange[i][2],
       ilosc: towaryDataRange[i][3],
       jednostka: towaryDataRange[i][4]
     });
  }

  // Ustandaryzowana nazwa firmy użytkownika do filtrowania
  const userFirmaSafe = userFirma ? String(userFirma).toLowerCase().trim() : "";

  for (let i = 1; i < awizDataRange.length; i++) {
    let row = awizDataRange[i];
    let idAwizacji = row[0];
    let rowUserId = row[1];
    
    let kontrahent = row[8] ? String(row[8]) : String(row[2]); 
    
    // Brak warunku IF - przepuszczamy każdy wpis prosto do tabeli
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
  
  
  // Sortowanie chronologiczne od najbliższych dat
  awizacje.sort(function(a, b) {
    return new Date(a.data) - new Date(b.data);
  });
  
  return awizacje;
}
