const url = 'https://script.google.com/macros/s/AKfycbzh8dvSWSPw7UqhJCU0-xsUs_aFZwAN2ytzVUW_19wwHKLUdc6BEFefnRckVmfoDI-aOA/exec?action=getData';

fetch(url)
  .then(res => res.json())
  .then(data => {
    console.log("=== RAW DATA FROM SHEET ===");
    console.log("CANDIDATES[0]:", data.candidates[0]);
    console.log("VOTES[0]:", data.votes[0]);

    // Map candidates like in App.jsx
    const mappedCandidates = data.candidates.map(c => {
      const id = String(c.id || c.ID || c.Id || '').trim();
      return {
        id: id,
        firstName: c.firstName || c.Nombre || c.nombre || '',
        lastName: c.lastName || c.Apellido || c.apellido || '',
      };
    });

    // Map votes like in App.jsx
    const mappedVotes = data.votes.map(v => {
      const candidateFirstName = v.candidateFirstName || v["Nombre Candidato"] || v.Nombre_Candidato || '';
      const candidateLastName = v.candidateLastName || v["Apellido Candidato"] || v.Apellido_Candidato || '';
      const voterFirstName = v.voterFirstName || v["Nombre Miembro"] || v.Nombre_Miembro || v.NombreMiembro || '';
      const voterLastName = v.voterLastName || v["Apellido Miembro"] || v.Apellido_Miembro || v.ApellidoMiembro || '';
      
      const oldCandName = v.candidateName || v.Nombre_Candidato || v["Nombre Candidato"] || v.Solicitante || v.solicitante || '';
      const oldVoterName = v.voterName || v.Nombre_Votante || v["Nombre Votante"] || v["Nombre del Miembro"] || '';

      const finalCandFirst = candidateFirstName || oldCandName.split(' ')[0] || '';
      const finalCandLast = candidateLastName || oldCandName.split(' ').slice(1).join(' ') || '';
      const finalVoterFirst = voterFirstName || oldVoterName.split(' ')[0] || '';
      const finalVoterLast = voterLastName || oldVoterName.split(' ').slice(1).join(' ') || '';

      return {
        candidateId: String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || v.candidateid || '').trim(),
        candidateFirstName: finalCandFirst,
        candidateLastName: finalCandLast,
        candidateName: `${finalCandFirst} ${finalCandLast}`.trim(),
        voterFirstName: finalVoterFirst,
        voterLastName: finalVoterLast,
        voterName: `${finalVoterFirst} ${finalVoterLast}`.trim(),
        status: v.status || v.Estado || v.estado || v.status || '',
        reason: v.reason || v.Motivo || v.motivo || v.reason || '',
      };
    });

    console.log("\n=== MAPPED DATA ===");
    console.log("MAPPED CANDIDATES:", mappedCandidates);
    console.log("MAPPED VOTES[0]:", mappedVotes[0]);

    // Test getCandidateStats like in AdminPanel.jsx
    const getCandidateStats = (candId, votesList) => {
      const candVotes = votesList.filter(v => {
        const vId = String(v.candidateId || v.ID_Candidato || v["ID Candidato"] || '').trim();
        const cId = String(candId || '').trim();
        const vIdInt = parseInt(vId, 10);
        const cIdInt = parseInt(cId, 10);
        return vId === cId || (!isNaN(vIdInt) && !isNaN(cIdInt) && vIdInt === cIdInt);
      });
      return {
        count: candVotes.length,
        votes: candVotes
      };
    };

    mappedCandidates.forEach(cand => {
      const stats = getCandidateStats(cand.id, mappedVotes);
      console.log(`\nCandidate: ${cand.firstName} ${cand.lastName} (ID: ${cand.id})`);
      console.log(`Matched votes count: ${stats.count}`);
      console.log("Matched votes details:", stats.votes);
    });
  })
  .catch(err => console.error(err));
