const GEOAPIFY_API_KEY = "4127fb8b9c3e418b8693ea38fdb4578b";
const MAX_DELAY_MINUTES = 60;
const MAX_DELAY_MS = MAX_DELAY_MINUTES * 60 * 1000;

const ideInput = document.getElementById("ideInput");
const dataFirstInput = document.getElementById("dataFirstInput");

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
  const file = ideInput.files[0];
  const dataFirstFile = dataFirstInput.files[0];
  const poiFile = poiInput.files[0];

  if (!file || !dataFirstFile || !poiFile) {
    alert("Please select ALL files.");
    return;
  }

  let ideData, dataFirstData;

  try {
    ideData = JSON.parse(await file.text());
    dataFirstData = JSON.parse(await dataFirstFile.text());
    POIS = JSON.parse(await poiFile.text());
  } catch {
    alert("Invalid JSON file.");
    return;
  }

  const fleet = [
    ...(ideData.fleet || []),
    ...(dataFirstData.data || [])
  ];

  if (!fleet.length) {
    alert("No fleet data found.");
    return;
  }

  const fleetGPS = await Promise.all(
    fleet.map(async (f) => {

      /* SUPPORT BOTH FORMATS */
      const Longitude = f.position?.LON ?? f.longitude;
      const Latitude = f.position?.LAT ?? f.latitude;

      const dateRaw = f.position?.date ?? f.dernier_signal;
      const date = new Date(dateRaw.replace(" ", "T"));
      const diffMs = Date.now() - date;

      let gpsStatus = "Outside POI";

      let locationParts = {
        city: null,
        client: null,
        commune: null,
        wilaya: null
      };

      /* GPS NOT UPDATED */
      if (diffMs > MAX_DELAY_MS) {
        return {
          Matricule: f.gps_alias ?? f.matricule,
          Agence: f.car_group ?? f.groupe_vehicule?.[0]?.thing_group_designation,
          Longitude,
          Latitude,
          Dernier_Date: dateRaw,
          Location: null,
          GPS_Status: `GPS non actualisé depuis ${dateRaw}`,
        };
      }

      /* POI CHECK */
      const matchedPOI = POIS.find(poi =>
        isPointInPolygon(Latitude, Longitude, poi.points)
      );

      if (matchedPOI) {
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
        Matricule: f.gps_alias ?? f.matricule,
        Agence: f.car_group ?? f.groupe_vehicule?.[0]?.thing_group_designation,
        Longitude,
        Latitude,
        Dernier_Date: dateRaw,
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
