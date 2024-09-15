// Fonction pour récupérer et afficher les payeurs depuis un fichier JSON
function fetchPayers(filename, elementId) {
    fetch(`data/${filename}`)
        .then(response => response.json())
        .then(data => {
            const ul = document.getElementById(elementId);
            data.forEach(payer => {
                const li = document.createElement('li');
                li.textContent = payer.email;  // Accède au champ email dans les objets JSON
                ul.appendChild(li);
            });
        })
        .catch(error => {
            console.error('Erreur lors du chargement des données :', error);
        });
}

// Appel de la fonction pour chaque liste de payeurs
fetchPayers('prompt_payers.json', 'good-payers');
fetchPayers('prompt_payers_without_form.json', 'late-payers');
fetchPayers('delinquent_payers.json', 'delinquent-payers');
