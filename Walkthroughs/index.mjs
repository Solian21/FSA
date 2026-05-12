import express from "express";
import path from "path";
import fs from "fs/promises";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;
const DATA_FILE = path.join(__dirname, "data.json");

app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

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

app.get("/", async (req, res) => {
  const data = await readData();

  res.render("home", {
    people: data.people,
    assignment: data.lastAssignment || null,
    error: null
  });
});

app.post("/toggle", async (req, res) => {
  const id = Number(req.body.id);
  const data = await readData();

  const person = data.people.find(person => person.id === id);

  if (person) {
    person.available = !person.available;
  }

  await writeData(data);
  res.redirect("/");
});

app.post("/generate", async (req, res) => {
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

app.post("/reset", async (req, res) => {
  const data = await readData();

  data.people.forEach(person => {
    person.promCount = 0;
    person.mainCount = 0;
  });

  data.lastAssignment = null;
  data.assignmentHistory = [];

  await writeData(data);

  res.redirect("/");
});

app.get("/team", async (req, res) => {
  const data = await readData();

  res.render("team", {
    people: data.people,
    error: null
  });
});

app.post("/team/add", async (req, res) => {
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

app.post("/team/delete", async (req, res) => {
  const data = await readData();
  const id = Number(req.body.id);

  data.people = data.people.filter(person => person.id !== id);

  await writeData(data);
  res.redirect("/team");
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});