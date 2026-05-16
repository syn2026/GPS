const GEOAPIFY_API_KEY = "4127fb8b9c3e418b8693ea38fdb4578b";

const MAX_DELAY_MINUTES = 60;
const MAX_DELAY_MS = MAX_DELAY_MINUTES * 60 * 1000;

/* INPUTS */
const ideInput = document.getElementById("ideInput");
const dataFirstInput = document.getElementById("dataFirstInput");
const poiInput = document.getElementById("poiInput");
const processBtn = document.getElementById("processBtn");

/* GLOBAL POI STORAGE */
let POIS = [];

/* =========================================
   POINT IN POLYGON
========================================= */
function isPointInPolygon(lat, lon, polygon) {
  let inside = false;

  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const xi = polygon[i][0];
    const yi = polygon[i][1];

    const xj = polygon[j][0];
    const yj = polygon[j][1];

    const intersect =
      yi > lat !== yj > lat &&
      lon < ((xj - xi) * (lat - yi)) / ((yj - yi) || 0.0000001) + xi;

    if (intersect) {
      inside = !inside;
    }
  }

  return inside;
}

/* =========================================
   SAFE DATE PARSER
========================================= */
function parseDate(dateRaw) {
  if (typeof dateRaw !== "string") {
    return null;
  }

  const clean = dateRaw.trim();

  if (!clean) {
    return null;
  }

  const date = new Date(clean.replace(" ", "T"));

  if (isNaN(date.getTime())) {
    return null;
  }

  return date;
}

/* =========================================
   REVERSE GEOCODING
========================================= */
async function reverseGeocode(lat, lon) {
  try {
    const response = await fetch(
      `https://api.geoapify.com/v1/geocode/reverse?lat=${lat}&lon=${lon}&lang=fr&apiKey=${GEOAPIFY_API_KEY}`
    );

    const json = await response.json();

    const props = json?.features?.[0]?.properties;

    return (
      props?.city ||
      props?.county ||
      props?.state ||
      props?.formatted ||
      null
    );
  } catch (error) {
    console.error("Reverse geocoding error:", error);
    return null;
  }
}

/* =========================================
   PROCESS VEHICLE
========================================= */
async function processVehicle(f) {
  const longitude = Number(f.position?.LON ?? f.longitude);
  const latitude = Number(f.position?.LAT ?? f.latitude);

  const matricule =
    f.gps_alias ??
    f.matricule ??
    f.thing_matricule ??
    "N/A";

  const agence =
    f.car_group ??
    f.groupe_vehicule?.[0]?.thing_group_designation ??
    f.groupes?.map(g => g.groupe_designation).join(", ") ??
    "N/A";

  const dateRaw = f.position?.date ?? f.date_dernier_signal;

  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) {
    return {
      Matricule: matricule,
      Agence: agence,
      Longitude: longitude,
      Latitude: latitude,
      Dernier_Date: dateRaw ?? null,
      Location: null,
      GPS_Status: "Coordonnées GPS invalides",
    };
  }

  const parsedDate = parseDate(dateRaw);

  let gpsStatus = "GPS actif";
  let location = null;

  if (parsedDate) {
    const diffMs = Date.now() - parsedDate.getTime();

    if (diffMs > MAX_DELAY_MS) {
      return {
        Matricule: matricule,
        Agence: agence,
        Longitude: longitude,
        Latitude: latitude,
        Dernier_Date: dateRaw,
        Location: null,
        GPS_Status: `GPS non actualisé depuis ${dateRaw}`,
      };
    }
  }

  const matchedPOI = POIS.find((poi) =>
  Array.isArray(poi.points) &&
  isPointInPolygon(latitude, longitude, poi.points)
);


  if (matchedPOI) {
    gpsStatus = "GPS actif (dans POI)";

    location = [
      matchedPOI.name,
      matchedPOI.commune,
      matchedPOI.wilaya,
    ]
      .filter(Boolean)
      .join(", ");
  } else {
    gpsStatus = "GPS actif (hors POI)";

    const city = await reverseGeocode(latitude, longitude);

    location = city;
  }

  return {
    Matricule: matricule,
    Agence: agence,
    Longitude: longitude,
    Latitude: latitude,
    Dernier_Date: dateRaw ?? null,
    Location: location,
    GPS_Status: gpsStatus,
  };
}

/* =========================================
   MAIN PROCESS
========================================= */
processBtn.addEventListener("click", async () => {
  try {
    const ideFile = ideInput.files?.[0];
    const dataFirstFile = dataFirstInput.files?.[0];
    const poiFile = poiInput.files?.[0];

    if (!ideFile || !dataFirstFile || !poiFile) {
      alert("Please select all required files.");
      return;
    }

    const ideData = JSON.parse(await ideFile.text());
    const dataFirstData = JSON.parse(await dataFirstFile.text());

    POIS = JSON.parse(await poiFile.text());

    const fleet = [
      ...(ideData?.fleet || []),
      ...(dataFirstData?.data || []),
    ];

    if (!fleet.length) {
      alert("No fleet data found.");
      return;
    }

    processBtn.disabled = true;
    processBtn.textContent = "Processing...";

   const fleetGPS = [];

for (const f of fleet) {
  try {
    const result = await processVehicle(f);
    fleetGPS.push(result);
  } catch (err) {
    console.error("Error processing vehicle:", err);
  }
}
``

    const worksheet = XLSX.utils.json_to_sheet(fleetGPS);
    const workbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      "FleetGPS"
    );

    XLSX.writeFile(workbook, "fleetGPS.xlsx");
    console.log("FINAL RESULT:", fleetGPS);

    alert("Excel exported successfully.");
  } catch (error) {
    console.error(error);
    alert("An error occurred while processing files.");
  } finally {
    processBtn.disabled = false;
    processBtn.textContent = "Process";
  }
});
