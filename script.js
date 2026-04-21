const GEOAPIFY_API_KEY = "4127fb8b9c3e418b8693ea38fdb4578b";
const MAX_DELAY_MINUTES = 60;
const MAX_DELAY_MS = MAX_DELAY_MINUTES * 60 * 1000;

const fileInput = document.getElementById("fileInput");
const poiInput = document.getElementById("poiInput");
const processBtn = document.getElementById("processBtn");

let POIS = [];

/* POINT IN POLYGON */
function isPointInPolygon(lat, lon, polygon) {
  let inside = false;

  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    const xi = polygon[i][0], yi = polygon[i][1];
    const xj = polygon[j][0], yj = polygon[j][1];

    const intersect =
      yi > lat !== yj > lat &&
      lon < ((xj - xi) * (lat - yi)) / (yj - yi + 0.0000001) + xi;

    if (intersect) inside = !inside;
  }

  return inside;
}

processBtn.addEventListener("click", async () => {
  const file = fileInput.files[0];
  const poiFile = poiInput.files[0];

  if (!file || !poiFile) {
    alert("Please select BOTH Fleet and POI files.");
    return;
  }

  let data;

  try {
    data = JSON.parse(await file.text());
    POIS = JSON.parse(await poiFile.text());
  } catch {
    alert("Invalid JSON file.");
    return;
  }

  if (!data || !Array.isArray(data.fleet)) {
    alert("JSON does not contain 'fleet' array.");
    return;
  }

  const fleetGPS = await Promise.all(
    data.fleet.map(async (f) => {
      const Longitude = f.position.LON;
      const Latitude = f.position.LAT;

      const date = new Date(f.position.date.replace(" ", "T"));
      const diffMs = Date.now() - date;

      let gpsStatus = "Outside POI";

      let locationParts = {
        city: null,
        client: null,
        commune: null,
        wilaya: null
      };

      /* ❌ GPS NOT UPDATED */
      if (diffMs > MAX_DELAY_MS) {
        return {
          Matricule: f.gps_alias,
          Agence: f.car_group,
          Longitude,
          Latitude,
          Dernier_Date: f.position.date,
          Location: null,
          GPS_Status: `GPS non actualisé depuis ${f.position.date}`,
        };
      }

      /* POI CHECK */
      const matchedPOI = POIS.find(poi =>
        isPointInPolygon(Latitude, Longitude, poi.points)
      );

      const insidePOI = !!matchedPOI;

      if (insidePOI) {
        gpsStatus = "GPS actif (dans POI)";

        locationParts.client = matchedPOI.name || null;
        locationParts.commune = matchedPOI.commune || null;
        locationParts.wilaya = matchedPOI.wilaya || null;

      } else {
        gpsStatus = "GPS actif (hors POI)";

        try {
          const res = await fetch(
            `https://api.geoapify.com/v1/geocode/reverse?lat=${Latitude}&lon=${Longitude}&lang=fr&apiKey=${GEOAPIFY_API_KEY}`
          );

          const json = await res.json();
          const props = json.features?.[0]?.properties;

          locationParts.city =
            props?.city ||
            props?.county ||
            props?.state ||
            null;

        } catch {}
      }

      /* CLEAN LOCATION FORMAT (NO EMPTY VALUES) */
      const Location = [
        locationParts.city,
        locationParts.client,
        locationParts.commune,
        locationParts.wilaya
      ]
        .map(v => (v ? v.toString().trim() : null))
        .filter(Boolean)
        .join(",");

      return {
        Matricule: f.gps_alias,
        Agence: f.car_group,
        Longitude,
        Latitude,
        Dernier_Date: f.position.date,
        Location,
        GPS_Status: gpsStatus,
      };
    })
  );

  const ws = XLSX.utils.json_to_sheet(fleetGPS);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "FleetGPS");

  XLSX.writeFile(wb, "fleetGPS.xlsx");

  alert("Done");
});
