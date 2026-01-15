async function isSharePointSiteReachable(siteUrl) {
    try {
        // Verwende die SharePoint REST-API über den aktuellen Kontext
        const response = await fetch(`${siteUrl}/_api/web`, {
            method: 'GET',
            headers: {
                'Accept': 'application/json;odata=verbose',
                // In SPFx wird der Auth-Token automatisch hinzugefügt
            },
            credentials: 'include' // Wichtig für Authentifizierung
        });

        if (response.ok) {
            console.log('✅ SharePoint-Site ist erreichbar:', siteUrl);
            return true;
        } else {
            console.warn('❌ SharePoint-Site ist nicht erreichbar (HTTP ' + response.status + '):', siteUrl);
            return false;
        }
    } catch (error) {
        console.error('❌ Fehler beim Prüfen der SharePoint-Site:', error.message);
        return false;
    }
}

// Beispielaufruf
isSharePointSiteReachable('https://ustgber.ecm.bundeswehr.org/arbeitsbereiche/A5620202414')
    .then(reachable => {
        if (reachable) {
            alert('Die SharePoint-Site ist erreichbar!');
        } else {
            alert('Die SharePoint-Site ist nicht erreichbar.');
        }
    });