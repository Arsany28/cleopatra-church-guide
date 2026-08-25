// ==========================================================================
// ✝️ كنيسة السيدة العذراء والبابا كيرلس عمود الدين - كليوباترا، الإسكندرية
// كود Google Apps Script لربط تطبيق الكنيسة بـ Google Sheets كقاعدة بيانات سحابية مجانية 100%
// ==========================================================================

// 1. دالة تهيئة الشيتات بالبيانات الافتراضية (Setup Database)
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet 1: Settings
  let shSettings = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  shSettings.clear();
  shSettings.appendRow(["Key", "Value"]);
  shSettings.appendRow(["churchName", "كنيسة السيدة العذراء والبابا كيرلس عمود الدين"]);
  shSettings.appendRow(["churchLocation", "📍 كليوباترا - الإسكندرية"]);
  shSettings.appendRow(["address", "شارع مصطفى باشا، كليوباترا حمامات، الإسكندرية"]);
  shSettings.appendRow(["shopWhatsapp", "201234567892"]);
  shSettings.appendRow(["studioPhone", "01234567890"]);
  shSettings.appendRow(["liveTitle", "بث صلوات القداس الإلهي والعشية مباشر"]);
  shSettings.appendRow(["youtubeUrl", "https://www.youtube-nocookie.com/embed/live_stream?channel=UC_SAMPLE"]);

  // Sheet 2: Places
  let shPlaces = ss.getSheetByName("Places") || ss.insertSheet("Places");
  shPlaces.clear();
  shPlaces.appendRow(["Key", "Title", "Floor", "Icon", "Description", "Hours", "Servant", "Phone", "Directions"]);
  shPlaces.appendRow(["altar", "الهيكل الرئيسي وحامل الأيقونات", "ground", "✝️", "المكان الأقدس حيث تقام صلوات القداس الإلهي وسر التناول.", "يومياً خلال مواعيد القداسات والعشيات", "آباء وخدام الكنيسة", "01234567890", "توجه إلى الأمام مباشرة من الباب الرئيسي للكنيسة نحو الشرق."]);
  shPlaces.appendRow(["bookstore", "المكتبة الكنسية والهدايا", "ground", "📖", "توفر أحدث الكتب الروحية، الأيقونات، الصلبان، وهدايا مدارس الأحد.", "يومياً 9ص-1م و 5م-9:30م", "خدام المكتبة", "01234567892", "عند الدخول من الباب الرئيسي، اتجه للجناح الأيمن بجوار المعرض."]);
  shPlaces.appendRow(["welcome", "مكتب استقبال وإرشاد الزوار", "ground", "ℹ️", "مخصص للترحيب بالزوار الجدد وتقديم الإرشادات والمساعدة.", "متاح طوال اليوم", "فريق خدمة الاستقبال", "01234567890", "موجود مباشرة عند المدخل الرئيسي للكنيسة."]);
  shPlaces.appendRow(["baptistery", "المعمودية ومزار القديسين", "ground", "🕊️", "مقر سر المعمودية المقدس ومزار للتبرك وإيقاد الشموع والصلاة.", "مفتوح للزيارة والصلاة طوال اليوم", "مسؤول المزار", "01234567890", "عند الدخول من الباب الرئيسي، اتجه للجناح الأيسر."]);
  shPlaces.appendRow(["sundayschool", "فصول مدارس الأحد والتعليم الكنسي", "first", "📚", "فصول مجهزة لشرح الكتاب المقدس والألحان والتاريخ الكنسي.", "الجمعة 11:00 ص - 1:30 م", "أمين الخدمة", "01234567890", "اصعد السلم الرئيسي إلى الدور الأول، الفصول موزعة على الممر."]);
  shPlaces.appendRow(["gameroom", "صالة الألعاب والأنشطة الرياضية", "first", "🏓", "تنس طاولة، بلياردو، بيبي فوت، وألعاب ذكاء جماعية للشباب.", "الجمعة والأحد بعد القداسات والاجتماعات", "مسؤول النشاط الرياضي", "01234567890", "الدور الأول - الجناح الأيمن بعد نهاية ممر الفصول."]);
  shPlaces.appendRow(["cafeteria", "كافتيريا الكنيسة", "basement", "☕", "مشروبات ساخنة وباردة ومأكولات صيامية وفطار طازج بعد القداسات.", "الجمعة والأحد والقداسات الرسمية", "فريق ضيافة الكافتيريا", "01234567890", "انزل سلم البدروم من الفناء الخارجي على اليمين."]);
  shPlaces.appendRow(["studio", "استوديو التسجيل والإنتاج الإعلامي", "basement", "🎙️", "استوديو صوتي مجهز بأحدث مايكروفونات وعزل احترافي للترانيم والإنتاج.", "بالحجز المسبق", "مهندس الصوت والإنتاج", "01234567890", "البدروم - الممر الداخلي المعزول صوتياً."]);

  // Sheet 3: Schedule
  let shSched = ss.getSheetByName("Schedule") || ss.insertSheet("Schedule");
  shSched.clear();
  shSched.appendRow(["ID", "Title", "Category", "Day", "Time", "Place", "Priest", "Badge"]);
  shSched.appendRow([1, "القداس الإلهي الأول (مبكر)", "liturgy", "الجمعة", "6:00 ص - 8:30 ص", "المذبح الرئيسي", "آباء الكنيسة", "قداس مبكر"]);
  shSched.appendRow([2, "القداس الإلهي الثاني (متأخر)", "liturgy", "الجمعة", "8:30 ص - 11:00 ص", "الكنيسة الكبرى", "آباء الكنيسة", "قداس متأخر"]);
  shSched.appendRow([3, "قداس يوم الأحد", "liturgy", "الأحد", "6:30 ص - 9:30 ص", "المذبح الرئيسي", "آباء الكنيسة", "أسبوعي"]);
  shSched.appendRow([4, "مدارس أحد - مرحلة الحضانة", "primary", "الجمعة", "11:15 ص - 12:30 م", "فصول الدور الأول", "خدام مرحلة الحضانة", "أطفال"]);
  shSched.appendRow([5, "مدارس أحد - المرحلة الابتدائية", "primary", "الجمعة", "11:15 ص - 1:00 م", "فصول الدور الأول", "خدام مرحلة ابتدائي", "ابتدائي"]);
  shSched.appendRow([6, "خدمة المرحلة الإعدادية", "prep-sec", "الجمعة", "11:30 ص - 1:15 م", "قاعة إعدادي - الدور الأول", "خدام مرحلة إعدادي", "إعدادي"]);
  shSched.appendRow([7, "خدمة المرحلة الثانوية", "prep-sec", "الجمعة", "6:30 م - 8:30 م", "قاعة الأنشطة", "خدام مرحلة ثانوي", "ثانوي"]);
  shSched.appendRow([8, "اجتماع الشباب والجامعيين", "youth", "الخميس", "7:00 م - 9:00 م", "قاعة الاجتماعات الكبرى", "مسؤول خدمة الشباب", "شباب"]);

  // Sheet 4: Products
  let shProd = ss.getSheetByName("Products") || ss.insertSheet("Products");
  shProd.clear();
  shProd.appendRow(["ID", "Name", "Category", "Price", "Image", "Badge"]);
  shProd.appendRow([101, "سيرة ودفاعيات البابا كيرلس عمود الدين", "cyril", 45, "📖", "الأكثر طلباً"]);
  shProd.appendRow([102, "أيقونة خشبية مذهبة للعذراء والبابا كيرلس", "icons", 120, "🖼️", "مباركة"]);
  shProd.appendRow([103, "كتاب صلوات الأجبية والقداس الإلهي (معرب)", "books", 35, "📕", "للمبتدئين"]);
  shProd.appendRow([104, "صليب خشب زيتون يدوي مبارك", "icons", 65, "✝️", "يدوي"]);
  shProd.appendRow([105, "كتاب تلوين وقصص الكتاب المقدس للأطفال", "kids", 30, "🎨", "أطفال"]);

  // Sheet 5: Clergy
  let shClergy = ss.getSheetByName("Clergy") || ss.insertSheet("Clergy");
  shClergy.clear();
  shClergy.appendRow(["ID", "Name", "Role", "Hours", "Place"]);
  shClergy.appendRow([1, "القمص / متياس", "كاهن الكنيسة", "الأربعاء والجمعة مساءً", "مكتب الكهنة - الدور الأرضي"]);
  shClergy.appendRow([2, "القس / أنطونيوس", "مسؤول خدمة الشباب", "الثلاثاء والسبت بعد العشية", "قاعة الشباب - الدور الأول"]);
  shClergy.appendRow([3, "القس / كيرلس", "مسؤول مدارس الأحد", "الأحد والخميس صباحاً ومساءً", "مكتب الكهنة - الدور الأرضي"]);

  // Format headers
  const sheets = [shSettings, shPlaces, shSched, shProd, shClergy];
  sheets.forEach(sh => {
    let rng = sh.getRange(1, 1, 1, sh.getLastColumn());
    rng.setBackground("#102A45").setFontColor("#FFFFFF").setFontWeight("bold");
    sh.setFrozenRows(1);
  });
}

// 2. دالة القراءة المباشرة من الشيت (GET API)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Settings
  let settings = {};
  let shSettings = ss.getSheetByName("Settings");
  if (shSettings) {
    let rows = shSettings.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0]) settings[rows[i][0]] = rows[i][1];
    }
  }

  // Places
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
          desc: String(rows[i][4] || ""),
          hours: String(rows[i][5] || ""),
          servant: String(rows[i][6] || ""),
          phone: String(rows[i][7] || ""),
          directions: String(rows[i][8] || "")
        });
      }
    }
  }

  // Schedule
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

  // Products
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

  // Clergy
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

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 3. دالة الحفظ والتعديل في الشيت (POST API)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
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

    // 2. Update Places
    if (data.places && Array.isArray(data.places)) {
      let sh = ss.getSheetByName("Places") || ss.insertSheet("Places");
      sh.clear();
      sh.appendRow(["Key", "Title", "Floor", "Icon", "Description", "Hours", "Servant", "Phone", "Directions"]);
      data.places.forEach(p => {
        sh.appendRow([p.key, p.title, p.floor, p.icon, p.desc, p.hours, p.servant, p.phone, p.directions]);
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

    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Google Sheets database updated successfully!" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
