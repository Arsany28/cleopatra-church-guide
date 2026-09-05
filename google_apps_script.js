// ==========================================================================
// ✝️ كنيسة السيدة العذراء والبابا كيرلس عمود الدين - كليوباترا، الإسكندرية
// كود Google Apps Script المتطور لدعم رفع وحفظ الصور الحقيقية للأماكن وكل الوسائط
// ==========================================================================

function doGet(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const fullDb = props.getProperty("FULL_DB_JSON");
    if (fullDb) {
      return ContentService.createTextOutput(fullDb).setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {}

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settings = {};
  let shSettings = ss.getSheetByName("Settings");
  if (shSettings) {
    let rows = shSettings.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
    }
  }

  let places = [];
  let shPlaces = ss.getSheetByName("Places");
  if (shPlaces) {
    let rows = shPlaces.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) {
        places.push({
          key: String(rows[i][0]),
          title: String(rows[i][1] || ""),
          floor: String(rows[i][2] || "ground"),
          icon: String(rows[i][3] || "🏛️"),
          image: String(rows[i][4] || ""),
          desc: String(rows[i][5] || ""),
          hours: String(rows[i][6] || ""),
          servant: String(rows[i][7] || ""),
          phone: String(rows[i][8] || ""),
          directions: String(rows[i][9] || "")
        });
      }
    }
  }

  let schedule = [];
  let shSched = ss.getSheetByName("Schedule");
  if (shSched) {
    let rows = shSched.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== "") {
        schedule.push({
          id: rows[i][0],
          title: String(rows[i][1] || ""),
          category: String(rows[i][2] || "liturgy"),
          day: String(rows[i][3] || ""),
          time: String(rows[i][4] || ""),
          place: String(rows[i][5] || ""),
          priest: String(rows[i][6] || ""),
          badge: String(rows[i][7] || "")
        });
      }
    }
  }

  let products = [];
  let shProd = ss.getSheetByName("Products");
  if (shProd) {
    let rows = shProd.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== "") {
        products.push({
          id: rows[i][0],
          name: String(rows[i][1] || ""),
          category: String(rows[i][2] || "books"),
          price: Number(rows[i][3] || 0),
          image: String(rows[i][4] || "📖"),
          badge: String(rows[i][5] || "")
        });
      }
    }
  }

  let clergy = [];
  let shClergy = ss.getSheetByName("Clergy");
  if (shClergy) {
    let rows = shClergy.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] !== "") {
        clergy.push({
          id: rows[i][0],
          name: String(rows[i][1] || ""),
          role: String(rows[i][2] || ""),
          hours: String(rows[i][3] || ""),
          place: String(rows[i][4] || "")
        });
      }
    }
  }

  const result = {
    settings: settings,
    places: places,
    schedule: schedule,
    products: products,
    clergy: clergy
  };

  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e.postData.contents;
    const data = JSON.parse(raw);
    
    // Save lossless complete JSON in ScriptProperties (Supports HD images & all places)
    try {
      PropertiesService.getScriptProperties().setProperty("FULL_DB_JSON", raw);
    } catch(err) {}

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Update Settings
    if (data.settings) {
      let sh = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
      sh.clear();
      sh.appendRow(["Key", "Value"]);
      for (let k in data.settings) {
        sh.appendRow([k, data.settings[k]]);
      }
    }

    // 2. Update Places (WITH REAL IMAGES)
    if (data.places && Array.isArray(data.places)) {
      let sh = ss.getSheetByName("Places") || ss.insertSheet("Places");
      sh.clear();
      sh.appendRow(["Key", "Title", "Floor", "Icon", "Image", "Description", "Hours", "Servant", "Phone", "Directions"]);
      data.places.forEach(p => {
        sh.appendRow([p.key, p.title, p.floor, p.icon, p.image || "", p.desc, p.hours, p.servant, p.phone, p.directions]);
      });
    }

    // 3. Update Schedule
    if (data.schedule && Array.isArray(data.schedule)) {
      let sh = ss.getSheetByName("Schedule") || ss.insertSheet("Schedule");
      sh.clear();
      sh.appendRow(["ID", "Title", "Category", "Day", "Time", "Place", "Priest", "Badge"]);
      data.schedule.forEach(s => {
        sh.appendRow([s.id, s.title, s.category, s.day, s.time, s.place, s.priest, s.badge]);
      });
    }

    // 4. Update Products
    if (data.products && Array.isArray(data.products)) {
      let sh = ss.getSheetByName("Products") || ss.insertSheet("Products");
      sh.clear();
      sh.appendRow(["ID", "Name", "Category", "Price", "Image", "Badge"]);
      data.products.forEach(pr => {
        sh.appendRow([pr.id, pr.name, pr.category, pr.price, pr.image, pr.badge]);
      });
    }

    // 5. Update Clergy
    if (data.clergy && Array.isArray(data.clergy)) {
      let sh = ss.getSheetByName("Clergy") || ss.insertSheet("Clergy");
      sh.clear();
      sh.appendRow(["ID", "Name", "Role", "Hours", "Place"]);
      data.clergy.forEach(c => {
        sh.appendRow([c.id, c.name, c.role, c.hours, c.place]);
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Google Sheets database updated successfully with images!" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
