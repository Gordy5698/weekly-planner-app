// ============================================================
// WEEKLY PLANNER → GOOGLE CALENDAR AUTO-SYNC
// ============================================================
// HOW TO USE:
//   1. Go to https://script.google.com
//   2. Click "New Project"
//   3. Delete any existing code and paste THIS entire file
//   4. Click the floppy disk icon to save (name it anything)
//   5. Click Run ▶️
//   6. Google will ask for permission — click Allow
//   7. Done! All events appear in your Google Calendar.
//
// To delete all synced events later, run: deleteSyncedEvents()
// ============================================================

var CALENDAR_NAME = "My Schedule"; // Change to your calendar name, or "primary" for main calendar
var TAG = "[planner-sync]"; // Used to track & delete events later

function syncToCalendar() {
  var cal = getOrCreateCalendar(CALENDAR_NAME);
  var events = getSchedule();
  var created = 0;

  events.forEach(function(e) {
    // Skip if already synced (check by title + date)
    var existing = cal.getEvents(e.start, e.end, { search: e.title });
    var alreadyExists = existing.some(function(ev) {
      return ev.getTitle() === e.title &&
             ev.getStartTime().getTime() === e.start.getTime();
    });

    if (!alreadyExists) {
      var event;
      if (e.allDay) {
        event = cal.createAllDayEvent(e.title, e.start, { description: e.desc });
      } else {
        event = cal.createEvent(e.title, e.start, e.end, { description: e.desc });
      }
      // Tag so we can clean up later
      event.setDescription((e.desc || "") + "\n\n" + TAG);
      created++;
    }
  });

  Logger.log("Done! Created " + created + " events in '" + CALENDAR_NAME + "'");
  SpreadsheetApp && SpreadsheetApp.getUi &&
    Browser.msgBox("Sync complete! Created " + created + " events.");
}

function deleteSyncedEvents() {
  var cal = getOrCreateCalendar(CALENDAR_NAME);
  var start = new Date("2026-06-28");
  var end   = new Date("2026-07-27");
  var events = cal.getEvents(start, end);
  var removed = 0;
  events.forEach(function(ev) {
    if ((ev.getDescription() || "").indexOf(TAG) !== -1) {
      ev.deleteEvent();
      removed++;
    }
  });
  Logger.log("Removed " + removed + " synced events.");
}

function getOrCreateCalendar(name) {
  if (name === "primary") return CalendarApp.getDefaultCalendar();
  var cals = CalendarApp.getCalendarsByName(name);
  if (cals.length > 0) return cals[0];
  return CalendarApp.createCalendar(name, { color: CalendarApp.Color.CYAN });
}

// ============================================================
// SCHEDULE DATA
// ============================================================
function getSchedule() {
  function dt(dateStr, timeStr) {
    // dateStr = "2026-06-28", timeStr = "08:00"
    var parts = timeStr.split(":");
    var d = new Date(dateStr);
    d.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
    return d;
  }
  function allDay(dateStr) {
    return new Date(dateStr);
  }

  return [
    // === SUN 28 JUN ===
    { title: "🏃 Morning Run (5k – easy pace)",
      start: dt("2026-06-28","08:00"), end: dt("2026-06-28","09:00"),
      desc: "Easy 5k to start the week. Keep HR low. Hydrate well after." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-06-28","12:30"), end: dt("2026-06-28","14:30"),
      desc: "New York market open. Focus on your watchlist. Journal every trade." },
    { title: "🍱 Meal Prep & Early Night",
      start: dt("2026-06-28","15:00"), end: dt("2026-06-28","18:00"),
      desc: "Prep meals for the week ahead. Aim for an early sleep — big week coming." },

    // === MON 29 JUN ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-06-29","09:00"), end: dt("2026-06-29","10:30"),
      desc: "Strength training. Focus on compound lifts. Don't max out — leave energy in the tank." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-06-29","12:30"), end: dt("2026-06-29","14:30"),
      desc: "New York market open. Stick to your plan. No revenge trading." },
    { title: "😴 Rest & Recovery",
      start: dt("2026-06-29","15:00"), end: dt("2026-06-29","17:00"),
      desc: "Post-gym recovery. Protein meal. Light stretching or foam roll." },

    // === TUE 30 JUN ===
    { title: "🏃 Morning Run (5–6k)",
      start: dt("2026-06-30","08:00"), end: dt("2026-06-30","09:00"),
      desc: "5–6k at conversational pace. Active recovery run after yesterday's gym." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-06-30","12:30"), end: dt("2026-06-30","14:30"),
      desc: "New York market open. Review your morning analysis before the session starts." },
    { title: "😴 Rest & Light Stretch",
      start: dt("2026-06-30","15:00"), end: dt("2026-06-30","16:00"),
      desc: "Light stretching or yoga. Eat a good dinner. Relax." },

    // === WED 1 JUL — LATE SHIFT ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-01","09:00"), end: dt("2026-07-01","10:30"),
      desc: "Strength training. Finish by 10:30 — work at 15:00. Eat a proper lunch after trading." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-01","12:30"), end: dt("2026-07-01","14:30"),
      desc: "Close positions and log trades by 14:30 — work at 15:00." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-01","15:00"), end: dt("2026-07-01","23:00"),
      desc: "Late shift. Eat well before you go in. Stay hydrated." },

    // === THU 2 JUL — LATE SHIFT ===
    { title: "🚶 Recovery Walk (Light Only)",
      start: dt("2026-07-02","09:00"), end: dt("2026-07-02","10:00"),
      desc: "Back-to-back late shifts. No gym or run today — short walk only. Food & hydration priority." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-02","12:30"), end: dt("2026-07-02","14:30"),
      desc: "Keep it simple today — back-to-back shifts. Don't overtrade." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-02","15:00"), end: dt("2026-07-02","23:00"),
      desc: "Late shift. Day 2 in a row — pace yourself. Eat well before going in." },

    // === FRI 3 JUL ===
    { title: "🏃 Morning Run (Easy – Post Shifts)",
      start: dt("2026-07-03","09:30"), end: dt("2026-07-03","10:30"),
      desc: "Keep it easy. You finished work at 23:00 last night. 4–5k at comfortable pace — no pressure." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-03","12:30"), end: dt("2026-07-03","14:30"),
      desc: "Free day to focus properly on the session." },
    { title: "😴 Rest Afternoon",
      start: dt("2026-07-03","15:00"), end: dt("2026-07-03","17:00"),
      desc: "Take it easy. Let your body recover from the back-to-back shifts." },

    // === SAT 4 JUL ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-04","09:00"), end: dt("2026-07-04","10:30"),
      desc: "Good solid session. No trading today — rest your mind too. Early Sunday shift tomorrow." },
    { title: "😴 Rest & Leisure – Recharge",
      start: dt("2026-07-04","12:00"), end: dt("2026-07-04","20:00"),
      desc: "No trading today. Rest and recharge. Early bed tonight — work starts at 07:00 tomorrow." },

    // === SUN 5 JUL — EARLY SHIFT ===
    { title: "💼 Work – Early Shift",
      start: dt("2026-07-05","07:00"), end: dt("2026-07-05","15:00"),
      desc: "Early shift. Set alarm for 06:00. Eat before you leave. No trading today — at work during NY open." },
    { title: "🚶 Light Walk & Stretch (Post Early Shift)",
      start: dt("2026-07-05","15:30"), end: dt("2026-07-05","16:30"),
      desc: "Gentle walk only after the early shift. No gym, no run. Let the body decompress. Early bed." },

    // === MON 6 JUL ===
    { title: "🏃 Morning Run (5k)",
      start: dt("2026-07-06","09:00"), end: dt("2026-07-06","10:00"),
      desc: "5k at easy pace. Good morning to shake off the early-shift hangover from yesterday." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-06","12:30"), end: dt("2026-07-06","14:30"),
      desc: "Full free afternoon — take the session seriously." },
    { title: "😴 Rest & Recovery",
      start: dt("2026-07-06","15:00"), end: dt("2026-07-06","17:00"),
      desc: "Rest up. Good meal in the evening." },

    // === TUE 7 JUL ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-07","09:00"), end: dt("2026-07-07","10:30"),
      desc: "Strength session. Push a bit harder today — free afternoon follows." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-07","12:30"), end: dt("2026-07-07","14:30"),
      desc: "Post-gym — eat well before sitting down to trade." },
    { title: "😴 Rest Afternoon",
      start: dt("2026-07-07","15:00"), end: dt("2026-07-07","17:00"),
      desc: "Chill. No obligations tonight." },

    // === WED 8 JUL — LATE SHIFT ===
    { title: "🏃 Morning Run (5k)",
      start: dt("2026-07-08","09:00"), end: dt("2026-07-08","10:00"),
      desc: "5k before a late shift. Running in the morning means you're done and full of energy for trading." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-08","12:30"), end: dt("2026-07-08","14:30"),
      desc: "Close out by 14:30 — shift starts at 15:00. Have lunch ready in advance." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-08","15:00"), end: dt("2026-07-08","23:00"),
      desc: "Late shift. Fuelled and ready after morning run and trading session." },

    // === THU 9 JUL ===
    { title: "🚶 Rest Morning – Post-Shift Recovery",
      start: dt("2026-07-09","09:00"), end: dt("2026-07-09","12:00"),
      desc: "Late finish last night at 23:00. Take the morning easy. Light walk at most." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-09","12:30"), end: dt("2026-07-09","14:30"),
      desc: "Post-shift — check you feel sharp enough before trading. Quality over quantity." },

    // === FRI 10 JUL ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-10","09:00"), end: dt("2026-07-10","10:30"),
      desc: "Good session today — fully free day. Push a solid workout." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-10","12:30"), end: dt("2026-07-10","14:30"),
      desc: "Free afternoon after — good day to really focus on the session." },
    { title: "😴 Rest – Prep for Tomorrow's Shift",
      start: dt("2026-07-10","15:00"), end: dt("2026-07-10","17:00"),
      desc: "Late shift tomorrow. Good night's sleep tonight. Don't stay up too late." },

    // === SAT 11 JUL — LATE SHIFT ===
    { title: "🚶 Light Morning Walk & Stretch",
      start: dt("2026-07-11","09:00"), end: dt("2026-07-11","10:00"),
      desc: "No hard training today — late shift ahead. Light walk and stretching only." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-11","12:30"), end: dt("2026-07-11","14:30"),
      desc: "Wrap up by 14:30 — shift in 30 mins. Lunch before you go in." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-11","15:00"), end: dt("2026-07-11","23:00"),
      desc: "Late shift." },

    // === SUN 12 JUL — LATE SHIFT ===
    { title: "😴 Rest – Back-to-Back Shifts",
      start: dt("2026-07-12","09:00"), end: dt("2026-07-12","12:00"),
      desc: "Second late shift in a row. Rest the body. Light walk only if you feel like it." },
    { title: "📊 Quick Trading Review (Short Session)",
      start: dt("2026-07-12","12:30"), end: dt("2026-07-12","14:00"),
      desc: "Shortened session — back-to-back shifts. Review charts and journal only. Don't force trades." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-12","15:00"), end: dt("2026-07-12","23:00"),
      desc: "Late shift. Stay fuelled — second day in a row." },

    // === MON 13 JUL ===
    { title: "🏃 Morning Run (Easy 5k – Post Shifts)",
      start: dt("2026-07-13","09:30"), end: dt("2026-07-13","10:30"),
      desc: "Easy pace only. You've done two late shifts — flush the legs out gently. 4–5k max." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-13","12:30"), end: dt("2026-07-13","14:30"),
      desc: "Free day — take your time with the session." },
    { title: "😴 Rest & Recovery",
      start: dt("2026-07-13","15:00"), end: dt("2026-07-13","17:00"),
      desc: "Recover properly. Protein meal. Light stretch." },

    // === TUE 14 JUL ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-14","09:00"), end: dt("2026-07-14","10:30"),
      desc: "Strength session. Body should be recovered now — good session today." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-14","12:30"), end: dt("2026-07-14","14:30"),
      desc: "Eat a good lunch between gym and the session." },
    { title: "😴 Rest Afternoon",
      start: dt("2026-07-14","15:00"), end: dt("2026-07-14","17:00"),
      desc: "Relax. No shift tonight." },

    // === WED 15 JUL — LATE SHIFT ===
    { title: "🏃 Morning Run (5–6k)",
      start: dt("2026-07-15","09:00"), end: dt("2026-07-15","10:00"),
      desc: "5–6k before a late shift. Solid routine: run in morning, trade at noon, work at 15:00." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-15","12:30"), end: dt("2026-07-15","14:30"),
      desc: "Cut off at 14:30 sharp — work at 15:00. Have your kit ready." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-15","15:00"), end: dt("2026-07-15","23:00"),
      desc: "Late shift." },

    // === THU 16 JUL ===
    { title: "🚶 Active Recovery Morning",
      start: dt("2026-07-16","09:00"), end: dt("2026-07-16","10:30"),
      desc: "Post-shift morning. Light walk or gentle stretch. No hard training — 3 late shifts incoming Fri/Sat/Sun." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-16","12:30"), end: dt("2026-07-16","14:30"),
      desc: "Free afternoon — sharp focus. Journal everything." },
    { title: "🎮 Leisure – Prep for 3-Day Stretch",
      start: dt("2026-07-16","15:00"), end: dt("2026-07-16","21:00"),
      desc: "Relax fully. Three late shifts in a row coming up. Food prep, early night." },

    // === FRI 17 JUL — LATE SHIFT ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-17","09:00"), end: dt("2026-07-17","10:30"),
      desc: "Strength session before the shift. Don't go too heavy — work at 15:00. Eat well after." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-17","12:30"), end: dt("2026-07-17","14:30"),
      desc: "Gym + trade + work — full productive day. Wrap up trading by 14:30." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-17","15:00"), end: dt("2026-07-17","23:00"),
      desc: "Late shift. Day 1 of 3 consecutive late shifts." },

    // === SAT 18 JUL — LATE SHIFT ===
    { title: "🚶 Light Morning Walk",
      start: dt("2026-07-18","09:00"), end: dt("2026-07-18","10:00"),
      desc: "Day 2 of 3 late shifts. No training today — light walk only. Conserve energy." },
    { title: "📊 Quick Trading Review",
      start: dt("2026-07-18","12:30"), end: dt("2026-07-18","14:00"),
      desc: "Shortened session. Back-to-back shifts — review and journal only if fatigued. Quality over quantity." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-18","15:00"), end: dt("2026-07-18","23:00"),
      desc: "Late shift. Day 2 of 3 consecutive late shifts." },

    // === SUN 19 JUL — LATE SHIFT + NIGHT STARTS ===
    { title: "🚶 Light Morning Walk",
      start: dt("2026-07-19","09:00"), end: dt("2026-07-19","10:00"),
      desc: "Day 3 of 3 late shifts + night shift starts tonight at 23:00. Light walk only. Rest and food." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-19","15:00"), end: dt("2026-07-19","23:00"),
      desc: "Late shift. Day 3 of 3. Night shift begins at 23:00 tonight — stay hydrated." },

    // === MON 20 JUL — NIGHT SHIFT ===
    { title: "😴 Rest & Night Shift Prep",
      start: dt("2026-07-20","09:00"), end: dt("2026-07-20","18:00"),
      desc: "Prepare your body for the night shift. Nap in the afternoon if possible. Eat a substantial meal before 22:00. No gym, no trading." },
    { title: "🌙 Work – Night Shift (23:00–07:00)",
      start: dt("2026-07-20","23:00"), end: dt("2026-07-21","07:00"),
      desc: "Night shift. Stay hydrated throughout. Bring food. Your body clock will be off — be kind to yourself." },

    // === TUE 21 JUL — POST NIGHT SHIFT RECOVERY ===
    { title: "😴 Sleep – Post Night Shift",
      start: dt("2026-07-21","07:00"), end: dt("2026-07-21","14:00"),
      desc: "Get home and sleep. Blackout curtains if you have them. No alarm — sleep as long as needed." },
    { title: "🚶 Very Light Walk & Recovery",
      start: dt("2026-07-21","15:00"), end: dt("2026-07-21","16:00"),
      desc: "Gentle 20-min walk. No gym, no run, no trading. Hydrate and eat good food. Sleep at a normal time tonight to reset your clock." },

    // === WED 22 JUL — GF GRADUATION ===
    { title: "🎓 GF Graduation – Liverpool Cathedral",
      start: allDay("2026-07-22"), end: allDay("2026-07-22"),
      allDay: true,
      desc: "Your girlfriend's graduation at Liverpool Cathedral. Travel to Liverpool. Celebrate properly — no training, no trading. This is her day. Enjoy every moment of it." },

    // === THU 23 JUL ===
    { title: "🏃 Morning Run – Gentle Return to Routine",
      start: dt("2026-07-23","09:00"), end: dt("2026-07-23","10:00"),
      desc: "Easy 4–5k. After night shift + graduation, ease back in gently. Movement will help reset your rhythm." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-23","12:30"), end: dt("2026-07-23","14:30"),
      desc: "Back to routine. Free afternoon — take your time and trade well." },
    { title: "😴 Rest Afternoon",
      start: dt("2026-07-23","15:00"), end: dt("2026-07-23","17:00"),
      desc: "Take it easy. Consistent sleep tonight for the late shifts ahead." },

    // === FRI 24 JUL — LATE SHIFT ===
    { title: "🏋️ Gym – Strength Session",
      start: dt("2026-07-24","09:00"), end: dt("2026-07-24","10:30"),
      desc: "Strength session before late shift. You're back in the groove — solid session." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-24","12:30"), end: dt("2026-07-24","14:30"),
      desc: "Gym + trade + work — strong day. Cut off at 14:30." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-24","15:00"), end: dt("2026-07-24","23:00"),
      desc: "Late shift." },

    // === SAT 25 JUL — LATE SHIFT ===
    { title: "🚶 Light Morning Walk",
      start: dt("2026-07-25","09:00"), end: dt("2026-07-25","10:00"),
      desc: "Back-to-back late shifts. Light walk only — no heavy training." },
    { title: "📊 Quick Trading Review",
      start: dt("2026-07-25","12:30"), end: dt("2026-07-25","14:00"),
      desc: "Short review session. Check open positions, journal yesterday's trades. Don't force new entries." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-25","15:00"), end: dt("2026-07-25","23:00"),
      desc: "Late shift." },

    // === SUN 26 JUL — LATE SHIFT ===
    { title: "🏃 Morning Run (Easy 5k)",
      start: dt("2026-07-26","09:00"), end: dt("2026-07-26","10:00"),
      desc: "Easy 5k to close out the stretch. You've worked hard — enjoy the run." },
    { title: "📈 Trading – NY Open Session",
      start: dt("2026-07-26","12:30"), end: dt("2026-07-26","14:30"),
      desc: "Last trading session of this block — make it count." },
    { title: "💼 Work – Late Shift",
      start: dt("2026-07-26","15:00"), end: dt("2026-07-26","23:00"),
      desc: "Late shift. Final shift of this planning block — well done for getting through it." }
  ];
}
