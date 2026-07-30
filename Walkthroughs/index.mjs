import "dotenv/config";
import express from "express";
import path from "path";
import fs from "fs/promises";
import { fileURLToPath } from "url";
import session from "express-session";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const DATA_FILE = path.join(__dirname, "data.json");
const TEMP_DATA_FILE = path.join(__dirname, "data.tmp.json");

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

app.set("trust proxy", 1);

app.use(
  session({
    secret: SESSION_SECRET,
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: process.env.NODE_ENV === "production",
      maxAge: 1000 * 60 * 60 * 12
    }
  })
);

/*
|--------------------------------------------------------------------------
| Middleware and helper functions
|--------------------------------------------------------------------------
*/

function requireLogin(req, res, next) {
  if (req.session.loggedIn) {
    return next();
  }

  return res.redirect("/login");
}

function asyncHandler(routeHandler) {
  return function wrappedRouteHandler(req, res, next) {
    Promise.resolve(routeHandler(req, res, next)).catch(next);
  };
}

async function readData() {
  const raw = await fs.readFile(DATA_FILE, "utf-8");
  const data = JSON.parse(raw);

  /*
   * Create default areas if this is an older data.json file.
   */
  if (!Array.isArray(data.areas)) {
    data.areas = [
      {
        id: 1,
        name: "Prom/NQ",
        peopleNeeded: 2,
        active: true
      },
      {
        id: 2,
        name: "Main",
        peopleNeeded: 2,
        active: true
      }
    ];
  }

  if (!Array.isArray(data.people)) {
    data.people = [];
  }

  if (!Array.isArray(data.assignmentHistory)) {
    data.assignmentHistory = [];
  }

  if (!Array.isArray(data.replacementHistory)) {
    data.replacementHistory = [];
  }

  if (data.lastAssignment === undefined) {
    data.lastAssignment = null;
  }

  /*
   * Normalize area data.
   */
  data.areas = data.areas
    .map(area => ({
      id: Number(area.id),
      name: String(area.name || "").trim(),
      peopleNeeded: Number(area.peopleNeeded),
      active: area.active !== false
    }))
    .filter(
      area =>
        Number.isInteger(area.id) &&
        area.id > 0 &&
        area.name &&
        Number.isInteger(area.peopleNeeded) &&
        area.peopleNeeded >= 1
    );

  /*
   * Normalize people and migrate old promCount/mainCount data.
   */
  data.people.forEach(person => {
    person.id = Number(person.id);
    person.name = String(person.name || "").trim();
    person.available = person.available !== false;

    if (
      !person.counts ||
      typeof person.counts !== "object" ||
      Array.isArray(person.counts)
    ) {
      person.counts = {};
    }

    const promArea = data.areas.find(area => area.name === "Prom/NQ");
    const mainArea = data.areas.find(area => area.name === "Main");

    if (
      promArea &&
      person.counts[promArea.id] === undefined &&
      Number.isInteger(person.promCount)
    ) {
      person.counts[promArea.id] = person.promCount;
    }

    if (
      mainArea &&
      person.counts[mainArea.id] === undefined &&
      Number.isInteger(person.mainCount)
    ) {
      person.counts[mainArea.id] = person.mainCount;
    }

    data.areas.forEach(area => {
      const count = Number(person.counts[area.id]);

      person.counts[area.id] =
        Number.isInteger(count) && count >= 0 ? count : 0;
    });

    delete person.promCount;
    delete person.mainCount;
  });

  /*
   * Convert an old Prom/Main assignment into the new groups format.
   */
  if (
    data.lastAssignment &&
    !Array.isArray(data.lastAssignment.groups) &&
    (Array.isArray(data.lastAssignment.prom) ||
      Array.isArray(data.lastAssignment.main))
  ) {
    const groups = [];

    const promArea = data.areas.find(area => area.name === "Prom/NQ");
    const mainArea = data.areas.find(area => area.name === "Main");

    if (promArea && Array.isArray(data.lastAssignment.prom)) {
      groups.push({
        areaId: promArea.id,
        areaName: promArea.name,
        people: data.lastAssignment.prom
      });
    }

    if (mainArea && Array.isArray(data.lastAssignment.main)) {
      groups.push({
        areaId: mainArea.id,
        areaName: mainArea.name,
        people: data.lastAssignment.main
      });
    }

    data.lastAssignment = {
      groups,
      date: data.lastAssignment.date || ""
    };
  }

  return data;
}

async function writeData(data) {
  const formattedData = JSON.stringify(data, null, 2);

  /*
   * Write to a temporary file first, then replace data.json.
   * This reduces the chance of data.json being left partially written.
   */
  await fs.writeFile(TEMP_DATA_FILE, formattedData, "utf-8");
  await fs.rename(TEMP_DATA_FILE, DATA_FILE);
}

function pickLowestRandom(people, areaId, amountNeeded) {
  const selected = [];

  while (selected.length < amountNeeded) {
    const remaining = people.filter(
      person => !selected.some(selectedPerson => selectedPerson.id === person.id)
    );

    if (remaining.length === 0) {
      break;
    }

    const lowestCount = Math.min(
      ...remaining.map(person => person.counts?.[areaId] ?? 0)
    );

    const tiedPeople = remaining.filter(
      person => (person.counts?.[areaId] ?? 0) === lowestCount
    );

    /*
     * Randomize people who have the same count.
     */
    tiedPeople.sort(() => Math.random() - 0.5);

    for (const person of tiedPeople) {
      if (selected.length >= amountNeeded) {
        break;
      }

      selected.push(person);
    }
  }

  return selected;
}

function getLastAssignmentIds(data) {
  return new Set(
    (data.lastAssignment?.groups || []).flatMap(group =>
      (group.people || []).map(person => Number(person.id))
    )
  );
}

function renderHome(res, data, options = {}) {
  return res.render("home", {
    people: data.people,
    areas: data.areas,
    assignment: options.assignment ?? data.lastAssignment ?? null,
    error: options.error ?? null
  });
}

function renderTeam(res, data, error = null) {
  return res.render("team", {
    people: data.people,
    areas: data.areas,
    error
  });
}

function renderFindSingle(
  res,
  data,
  {
    replacement = null,
    area = null,
    error = null
  } = {}
) {
  return res.render("find-single", {
    people: data.people,
    areas: data.areas.filter(areaItem => areaItem.active),
    replacement,
    area,
    error
  });
}

/*
|--------------------------------------------------------------------------
| Login routes
|--------------------------------------------------------------------------
*/

app.get("/login", (req, res) => {
  if (req.session.loggedIn) {
    return res.redirect("/");
  }

  return res.render("login", {
    error: null
  });
});

app.post("/login", (req, res) => {
  const password = req.body.password?.trim();

  if (password === SITE_PASSWORD) {
    req.session.loggedIn = true;
    return res.redirect("/");
  }

  return res.status(401).render("login", {
    error: "Incorrect password. Please try again."
  });
});

app.post("/logout", (req, res, next) => {
  req.session.destroy(error => {
    if (error) {
      return next(error);
    }

    res.clearCookie("connect.sid");
    return res.redirect("/login");
  });
});

/*
|--------------------------------------------------------------------------
| Home and assignment routes
|--------------------------------------------------------------------------
*/

app.get(
  "/",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    return renderHome(res, data);
  })
);

app.post(
  "/toggle",
  requireLogin,
  asyncHandler(async (req, res) => {
    const id = Number(req.body.id);
    const data = await readData();

    const person = data.people.find(personItem => personItem.id === id);

    if (person) {
      person.available = !person.available;
      await writeData(data);
    }

    return res.redirect("/");
  })
);

app.post(
  "/generate",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();

    const activeAreas = data.areas.filter(
      area => area.active && area.peopleNeeded > 0
    );

    if (activeAreas.length === 0) {
      return renderHome(res, data, {
        error:
          "There are no active areas. Go to Manage Areas and activate or add an area."
      });
    }

    const totalPeopleNeeded = activeAreas.reduce(
      (total, area) => total + area.peopleNeeded,
      0
    );

    const allAvailablePeople = data.people.filter(
      person => person.available
    );

    if (allAvailablePeople.length < totalPeopleNeeded) {
      return renderHome(res, data, {
        error:
          `The active areas require ${totalPeopleNeeded} people, ` +
          `but only ${allAvailablePeople.length} are available.`
      });
    }

    const previousIds = getLastAssignmentIds(data);

    let availablePeople = allAvailablePeople.filter(
      person => !previousIds.has(person.id)
    );

    /*
     * If there are not enough people after excluding the previous assignment,
     * allow those people back into the selection.
     */
    if (availablePeople.length < totalPeopleNeeded) {
      availablePeople = allAvailablePeople;
    }

    const selectedIds = new Set();
    const groups = [];

    for (const area of activeAreas) {
      const candidates = availablePeople.filter(
        person => !selectedIds.has(person.id)
      );

      const selectedPeople = pickLowestRandom(
        candidates,
        area.id,
        area.peopleNeeded
      );

      if (selectedPeople.length < area.peopleNeeded) {
        return renderHome(res, data, {
          error: `Not enough people could be selected for ${area.name}.`
        });
      }

      selectedPeople.forEach(person => {
        selectedIds.add(person.id);

        person.counts ??= {};
        person.counts[area.id] =
          (person.counts[area.id] ?? 0) + 1;
      });

      groups.push({
        areaId: area.id,
        areaName: area.name,
        people: selectedPeople.map(person => ({
          id: person.id,
          name: person.name
        }))
      });
    }

    const assignment = {
      groups,
      date: new Date().toLocaleDateString("en-US")
    };

    data.lastAssignment = assignment;
    data.assignmentHistory.push(assignment);

    await writeData(data);

    return renderHome(res, data, {
      assignment
    });
  })
);

/*
|--------------------------------------------------------------------------
| Single replacement routes
|--------------------------------------------------------------------------
*/

app.get(
  "/find-single",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    return renderFindSingle(res, data);
  })
);

app.post(
  "/find-single",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const areaId = Number(req.body.areaId);

    const selectedArea = data.areas.find(
      area => area.id === areaId && area.active
    );

    if (!selectedArea) {
      return renderFindSingle(res, data, {
        error: "Please select a valid active area."
      });
    }

    const lastAssignmentIds = getLastAssignmentIds(data);

    let availablePeople = data.people.filter(
      person =>
        person.available &&
        !lastAssignmentIds.has(person.id)
    );

    /*
     * If everybody available is already in the current assignment,
     * allow assigned people back into the candidate list.
     */
    if (availablePeople.length === 0) {
      availablePeople = data.people.filter(
        person => person.available
      );
    }

    if (availablePeople.length === 0) {
      return renderFindSingle(res, data, {
        area: selectedArea,
        error: "No available people were found."
      });
    }

    const replacementGroup = pickLowestRandom(
      availablePeople,
      selectedArea.id,
      1
    );

    const replacement = replacementGroup[0];

    if (!replacement) {
      return renderFindSingle(res, data, {
        area: selectedArea,
        error: "A replacement could not be selected."
      });
    }

    replacement.counts ??= {};
    replacement.counts[selectedArea.id] =
      (replacement.counts[selectedArea.id] ?? 0) + 1;

    const replacementRecord = {
      areaId: selectedArea.id,
      areaName: selectedArea.name,
      person: {
        id: replacement.id,
        name: replacement.name
      },
      date: new Date().toLocaleDateString("en-US")
    };

    data.replacementHistory.push(replacementRecord);

    await writeData(data);

    return renderFindSingle(res, data, {
      replacement: replacementRecord.person,
      area: selectedArea
    });
  })
);

/*
|--------------------------------------------------------------------------
| Reset route
|--------------------------------------------------------------------------
*/

app.post(
  "/reset",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();

    data.people.forEach(person => {
      person.counts = {};

      data.areas.forEach(area => {
        person.counts[area.id] = 0;
      });
    });

    data.lastAssignment = null;
    data.assignmentHistory = [];
    data.replacementHistory = [];

    await writeData(data);

    return res.redirect("/");
  })
);

/*
|--------------------------------------------------------------------------
| Team management routes
|--------------------------------------------------------------------------
*/

app.get(
  "/team",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    return renderTeam(res, data);
  })
);

app.post(
  "/team/add",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const name = req.body.name?.trim();

    if (!name) {
      return renderTeam(res, data, "Name cannot be empty.");
    }

    const duplicateName = data.people.some(
      person => person.name.toLowerCase() === name.toLowerCase()
    );

    if (duplicateName) {
      return renderTeam(
        res,
        data,
        "A team member with that name already exists."
      );
    }

    const newId =
      data.people.length > 0
        ? Math.max(...data.people.map(person => person.id)) + 1
        : 1;

    const counts = {};

    data.areas.forEach(area => {
      counts[area.id] = 0;
    });

    data.people.push({
      id: newId,
      name,
      available: true,
      counts
    });

    await writeData(data);

    return res.redirect("/team");
  })
);

app.post(
  "/team/delete",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.body.id);

    data.people = data.people.filter(person => person.id !== id);

    await writeData(data);

    return res.redirect("/team");
  })
);

app.get(
  "/team/edit/:id",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.params.id);

    const person = data.people.find(
      personItem => personItem.id === id
    );

    if (!person) {
      return res.redirect("/team");
    }

    return res.render("edit-person", {
      person,
      areas: data.areas,
      error: null
    });
  })
);

app.post(
  "/team/edit/:id",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.params.id);

    const person = data.people.find(
      personItem => personItem.id === id
    );

    if (!person) {
      return res.redirect("/team");
    }

    const name = req.body.name?.trim();
    const available = req.body.available === "on";

    if (!name) {
      return res.render("edit-person", {
        person,
        areas: data.areas,
        error: "Name cannot be empty."
      });
    }

    const duplicateName = data.people.some(
      otherPerson =>
        otherPerson.id !== id &&
        otherPerson.name.toLowerCase() === name.toLowerCase()
    );

    if (duplicateName) {
      return res.render("edit-person", {
        person,
        areas: data.areas,
        error: "A team member with that name already exists."
      });
    }

    const updatedCounts = {};

    for (const area of data.areas) {
      /*
       * The edit-person form should name each input:
       * count_<area id>
       *
       * Examples:
       * count_1
       * count_2
       */
      const submittedCount = Number(req.body[`count_${area.id}`]);

      if (
        !Number.isInteger(submittedCount) ||
        submittedCount < 0
      ) {
        return res.render("edit-person", {
          person,
          areas: data.areas,
          error:
            `The ${area.name} count must be a non-negative whole number.`
        });
      }

      updatedCounts[area.id] = submittedCount;
    }

    person.name = name;
    person.available = available;
    person.counts = updatedCounts;

    await writeData(data);

    return res.redirect("/team");
  })
);

/*
|--------------------------------------------------------------------------
| Area management routes
|--------------------------------------------------------------------------
*/

app.get(
  "/areas",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();

    return res.render("areas", {
      areas: data.areas,
      error: null
    });
  })
);

app.post(
  "/areas/add",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();

    const name = req.body.name?.trim();
    const peopleNeeded = Number(req.body.peopleNeeded);

    if (!name) {
      return res.render("areas", {
        areas: data.areas,
        error: "Area name cannot be empty."
      });
    }

    if (
      !Number.isInteger(peopleNeeded) ||
      peopleNeeded < 1
    ) {
      return res.render("areas", {
        areas: data.areas,
        error:
          "People needed must be a whole number of at least 1."
      });
    }

    const duplicate = data.areas.some(
      area => area.name.toLowerCase() === name.toLowerCase()
    );

    if (duplicate) {
      return res.render("areas", {
        areas: data.areas,
        error: "An area with that name already exists."
      });
    }

    const newId =
      data.areas.length > 0
        ? Math.max(...data.areas.map(area => area.id)) + 1
        : 1;

    data.areas.push({
      id: newId,
      name,
      peopleNeeded,
      active: true
    });

    data.people.forEach(person => {
      person.counts ??= {};
      person.counts[newId] = 0;
    });

    await writeData(data);

    return res.redirect("/areas");
  })
);

app.get(
  "/areas/edit/:id",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.params.id);

    const area = data.areas.find(
      areaItem => areaItem.id === id
    );

    if (!area) {
      return res.redirect("/areas");
    }

    return res.render("edit-area", {
      area,
      error: null
    });
  })
);

app.post(
  "/areas/edit/:id",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.params.id);

    const area = data.areas.find(
      areaItem => areaItem.id === id
    );

    if (!area) {
      return res.redirect("/areas");
    }

    const name = req.body.name?.trim();
    const peopleNeeded = Number(req.body.peopleNeeded);
    const active = req.body.active === "on";

    if (!name) {
      return res.render("edit-area", {
        area,
        error: "Area name cannot be empty."
      });
    }

    if (
      !Number.isInteger(peopleNeeded) ||
      peopleNeeded < 1
    ) {
      return res.render("edit-area", {
        area,
        error:
          "People needed must be a whole number of at least 1."
      });
    }

    const duplicate = data.areas.some(
      otherArea =>
        otherArea.id !== id &&
        otherArea.name.toLowerCase() === name.toLowerCase()
    );

    if (duplicate) {
      return res.render("edit-area", {
        area,
        error: "An area with that name already exists."
      });
    }

    area.name = name;
    area.peopleNeeded = peopleNeeded;
    area.active = active;

    /*
     * Update the displayed name in the current assignment.
     */
    if (Array.isArray(data.lastAssignment?.groups)) {
      const assignmentGroup = data.lastAssignment.groups.find(
        group => Number(group.areaId) === id
      );

      if (assignmentGroup) {
        assignmentGroup.areaName = name;
      }
    }

    await writeData(data);

    return res.redirect("/areas");
  })
);

app.post(
  "/areas/delete",
  requireLogin,
  asyncHandler(async (req, res) => {
    const data = await readData();
    const id = Number(req.body.id);

    const areaExists = data.areas.some(area => area.id === id);

    if (!areaExists) {
      return res.redirect("/areas");
    }

    data.areas = data.areas.filter(area => area.id !== id);

    data.people.forEach(person => {
      if (person.counts) {
        delete person.counts[id];
      }
    });

    /*
     * Remove the deleted area from the current assignment.
     * Historical assignments remain unchanged.
     */
    if (Array.isArray(data.lastAssignment?.groups)) {
      data.lastAssignment.groups =
        data.lastAssignment.groups.filter(
          group => Number(group.areaId) !== id
        );

      if (data.lastAssignment.groups.length === 0) {
        data.lastAssignment = null;
      }
    }

    await writeData(data);

    return res.redirect("/areas");
  })
);

/*
|--------------------------------------------------------------------------
| Not found and error handling
|--------------------------------------------------------------------------
*/

app.use((req, res) => {
  res.status(404).send("Page not found.");
});

app.use((error, req, res, next) => {
  console.error(error);

  if (res.headersSent) {
    return next(error);
  }

  return res.status(500).send(
    "Something went wrong. Check the server console for details."
  );
});

/*
|--------------------------------------------------------------------------
| Start server
|--------------------------------------------------------------------------
*/

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});