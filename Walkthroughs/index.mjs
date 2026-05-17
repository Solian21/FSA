import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs/promises";
import { fileURLToPath } from "url";
import session from "express-session";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, "data.json");

const SITE_PASSWORD = process.env.PASSWORD;
const SESSION_SECRET = process.env.SESSION_SECRET || "backup-secret";

if (!SITE_PASSWORD) {
  console.error("Missing PASSWORD in .env file");
  process.exit(1);
}

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false
  })
);

function requireLogin(req, res, next) {
  if (req.session.loggedIn) {
    return next();
  }

  res.redirect("/login");
}

async function readData() {
  const raw = await fs.readFile(DATA_FILE, "utf-8");
  return JSON.parse(raw);
}

async function writeData(data) {
  await fs.writeFile(DATA_FILE, JSON.stringify(data, null, 2));
}

function pickLowestRandom(people, countKey, amountNeeded) {
  const selected = [];

  while (selected.length < amountNeeded) {
    const remaining = people.filter(person => !selected.includes(person));

    if (remaining.length === 0) break;

    const lowestCount = Math.min(...remaining.map(person => person[countKey]));

    const tiedPeople = remaining.filter(
      person => person[countKey] === lowestCount
    );

    tiedPeople.sort(() => Math.random() - 0.5);

    selected.push(...tiedPeople);
  }

  return selected.slice(0, amountNeeded);
}

// Login page
app.get("/login", (req, res) => {
  if (req.session.loggedIn) {
    return res.redirect("/");
  }

  res.render("login", {
    error: null
  });
});

// Login form submit
app.post("/login", (req, res) => {
  const password = req.body.password?.trim();

  if (password === SITE_PASSWORD) {
    req.session.loggedIn = true;
    return res.redirect("/");
  }

  res.render("login", {
    error: "Incorrect password. Please try again."
  });
});

// Logout
app.post("/logout", (req, res) => {
  req.session.destroy(() => {
    res.redirect("/login");
  });
});

// Protected home page
app.get("/", requireLogin, async (req, res) => {
  const data = await readData();

  res.render("home", {
    people: data.people,
    assignment: data.lastAssignment || null,
    error: null
  });
});

app.post("/toggle", requireLogin, async (req, res) => {
  const id = Number(req.body.id);
  const data = await readData();

  const person = data.people.find(person => person.id === id);

  if (person) {
    person.available = !person.available;
  }

  await writeData(data);
  res.redirect("/");
});

app.post("/generate", requireLogin, async (req, res) => {
  const data = await readData();

  const lastAssignmentPeople = [
    ...(data.lastAssignment?.prom || []),
    ...(data.lastAssignment?.main || [])
  ];

  const lastAssignmentIds = lastAssignmentPeople.map(person => person.id);

  let availablePeople = data.people.filter(
    person => person.available && !lastAssignmentIds.includes(person.id)
  );

  // If avoiding last week's people leaves fewer than 4,
  // allow them back in because it is unavoidable.
  if (availablePeople.length < 4) {
    availablePeople = data.people.filter(person => person.available);
  }

  if (availablePeople.length < 4) {
    return res.render("home", {
      people: data.people,
      assignment: data.lastAssignment || null,
      error: "Not enough available people. You need at least 4 available people."
    });
  }

  const promGroup = pickLowestRandom(availablePeople, "promCount", 2);

  const remainingPeople = availablePeople.filter(
    person => !promGroup.includes(person)
  );

  const mainGroup = pickLowestRandom(remainingPeople, "mainCount", 2);

  promGroup.forEach(person => {
    person.promCount += 1;
  });

  mainGroup.forEach(person => {
    person.mainCount += 1;
  });

  const assignment = {
    prom: promGroup.map(person => ({
      id: person.id,
      name: person.name
    })),
    main: mainGroup.map(person => ({
      id: person.id,
      name: person.name
    })),
    date: new Date().toLocaleDateString()
  };

  data.lastAssignment = assignment;

  if (!data.assignmentHistory) {
    data.assignmentHistory = [];
  }

  data.assignmentHistory.push(assignment);

  await writeData(data);

  res.render("home", {
    people: data.people,
    assignment,
    error: null
  });
});

// Find single replacement page
app.get("/find-single", requireLogin, async (req, res) => {
  const data = await readData();

  res.render("find-single", {
    people: data.people,
    replacement: null,
    area: null,
    error: null
  });
});

// Generate one replacement for either prom or main
app.post("/find-single", requireLogin, async (req, res) => {
  const data = await readData();
  const area = req.body.area;

  if (area !== "prom" && area !== "main") {
    return res.render("find-single", {
      people: data.people,
      replacement: null,
      area: null,
      error: "Please choose either Prom or Main."
    });
  }

  const countKey = area === "prom" ? "promCount" : "mainCount";

  const lastAssignmentPeople = [
    ...(data.lastAssignment?.prom || []),
    ...(data.lastAssignment?.main || [])
  ];

  const lastAssignmentIds = lastAssignmentPeople.map(person => person.id);

  let availablePeople = data.people.filter(
    person => person.available && !lastAssignmentIds.includes(person.id)
  );

  // If avoiding the last assignment leaves nobody,
  // allow last assignment people back in because it is unavoidable.
  if (availablePeople.length < 1) {
    availablePeople = data.people.filter(person => person.available);
  }

  if (availablePeople.length < 1) {
    return res.render("find-single", {
      people: data.people,
      replacement: null,
      area,
      error: "No available people found."
    });
  }

  const replacementGroup = pickLowestRandom(availablePeople, countKey, 1);
  const replacement = replacementGroup[0];

  replacement[countKey] += 1;

  const replacementRecord = {
    area,
    person: {
      id: replacement.id,
      name: replacement.name
    },
    date: new Date().toLocaleDateString()
  };

  if (!data.replacementHistory) {
    data.replacementHistory = [];
  }

  data.replacementHistory.push(replacementRecord);

  await writeData(data);

  res.render("find-single", {
    people: data.people,
    replacement: replacementRecord.person,
    area,
    error: null
  });
});

app.post("/reset", requireLogin, async (req, res) => {
  const data = await readData();

  data.people.forEach(person => {
    person.promCount = 0;
    person.mainCount = 0;
  });

  data.lastAssignment = null;
  data.assignmentHistory = [];
  data.replacementHistory = [];

  await writeData(data);

  res.redirect("/");
});

app.get("/team", requireLogin, async (req, res) => {
  const data = await readData();

  res.render("team", {
    people: data.people,
    error: null
  });
});

app.post("/team/add", requireLogin, async (req, res) => {
  const data = await readData();
  const name = req.body.name.trim();

  if (!name) {
    return res.render("team", {
      people: data.people,
      error: "Name cannot be empty."
    });
  }

  const newId =
    data.people.length > 0
      ? Math.max(...data.people.map(person => person.id)) + 1
      : 1;

  data.people.push({
    id: newId,
    name,
    promCount: 0,
    mainCount: 0,
    available: true
  });

  await writeData(data);
  res.redirect("/team");
});

app.post("/team/delete", requireLogin, async (req, res) => {
  const data = await readData();
  const id = Number(req.body.id);

  data.people = data.people.filter(person => person.id !== id);

  await writeData(data);
  res.redirect("/team");
});

app.get("/team/edit/:id", requireLogin, async (req, res) => {
  const data = await readData();
  const id = Number(req.params.id);

  const person = data.people.find(person => person.id === id);

  if (!person) {
    return res.redirect("/team");
  }

  res.render("edit-person", {
    person,
    error: null
  });
});

app.post("/team/edit/:id", requireLogin, async (req, res) => {
  const data = await readData();
  const id = Number(req.params.id);

  const person = data.people.find(person => person.id === id);

  if (!person) {
    return res.redirect("/team");
  }

  const name = req.body.name?.trim();
  const promCount = Number(req.body.promCount);
  const mainCount = Number(req.body.mainCount);
  const available = req.body.available === "on";

  if (!name) {
    return res.render("edit-person", {
      person,
      error: "Name cannot be empty."
    });
  }

  if (promCount < 0 || mainCount < 0) {
    return res.render("edit-person", {
      person,
      error: "Counts cannot be negative."
    });
  }

  person.name = name;
  person.promCount = promCount;
  person.mainCount = mainCount;
  person.available = available;

  await writeData(data);

  res.redirect("/team");
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});