import axios from 'axios'; 

const config = {
    headers: {
        // 'Authorization': 'Bearer your-token-here',
        'Content-Type': 'application/json', 
    }
}; 

const data = {
    schlagwoerter: 'Kloepfel Consulting GmbH',
    schlagwortOptionen: '3', 
    niederlassung: '*', 
    registerNummer: '', 
    registerGericht: ''
}

async function fetchData() {
    try {
        const response = await axios.post('https://www.handelsregister.de/rp_web/erweitertesuche.xhtml', data, config);
        console.log(response);
    } catch (error) {
        console.error('Error fetching data:', error);
    }
}

fetchData();
