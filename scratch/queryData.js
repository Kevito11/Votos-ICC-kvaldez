const url = 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec?action=getData';

fetch(url)
  .then(res => res.json())
  .then(data => {
    console.log("CANDIDATES (first 3):");
    console.log(data.candidates.slice(0, 3));
    console.log("\nALL VOTES:");
    console.log(data.votes);
  })
  .catch(err => console.error(err));
