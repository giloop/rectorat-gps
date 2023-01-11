// Création de la map 
var map_0 = L.map("map_0", {
    center: [46.7687714, 4.5660859],
    crs: L.CRS.EPSG3857,
    zoom: 2.1,
    zoomControl: true,
    preferCanvas: false,
    id: "my_map",
});

var tile_layer_1 = L.tileLayer(
    "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        "attribution": "Data by \u0026copy; \u003ca href=\"http://openstreetmap.org\"\u003eOpenStreetMap\u003c/a\u003e, under \u003ca href=\"http://www.openstreetmap.org/copyright\"\u003eODbL\u003c/a\u003e.",
        "detectRetina": false,
        "maxNativeZoom": 18,
        "maxZoom": 18,
        "minZoom": 1,
        "noWrap": false,
        "opacity": 1,
        "subdomains": "abc",
        "tms": false
    }
).addTo(map_0);

// Création des sous-groupes 
var groups = {
    'Autre': null
}

var arKeys = Object.keys(groups)

// Markers de type 
var markers = L.markerClusterGroup({});
map_0.addLayer(markers);

// Ajout des sous-groupes à la map
arKeys.forEach(el => {
    groups[el] = L.featureGroup.subGroup(markers);
    groups[el].addTo(map_0);
})


function addMarker(item) {
    // Item : 
    // {"status": "etab",
    // "organisme": row['Organisme candidat'],
    // "coordinates" : coords,
    // "couleur": "lightgray" if row['CodePostal'] == "Non trouvé" else "lightred",
    // "infos": f"{info}""}

    foundKeys = arKeys.filter(el => item.organisme.includes(el))
    if (foundKeys.length > 0) {
        myGroup = foundKeys[0]
    } else {
        myGroup = 'Autre'
    }

    var marker = L.marker(item.coordinates, {})
        .addTo(groups[myGroup])
        .bindPopup(item.infos)
        .bindTooltip(item.infos, { "sticky": true })
        .setIcon(L.AwesomeMarkers.icon({
            "extraClasses": "fa-rotate-0",
            "icon": "ok-sign",
            "iconColor": "white",
            "markerColor": item.couleur,
            "prefix": "glyphicon"
        }));
}

// Chargement des données
fetch('carto_programmes.json')
    .then(function(response) {
        return response.json()
    })
    .then(function(json) {
        // traitement du JSON
        json.forEach(el => { addMarker(el) });
    })
    .catch(function(error) {
        console.log('Il y a eu un problème avec l\'opération fetch: ' + error.message);
    });




var myLayers = {
    base_layers: {
        "openstreetmap": tile_layer_1,
    },
    overlays: groups,
};

L.control.layers(
    myLayers.base_layers,
    myLayers.overlays, {
        "autoZIndex": true,
        "collapsed": true,
        "position": "topright"
    }
).addTo(map_0);