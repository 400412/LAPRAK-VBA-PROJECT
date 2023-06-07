// 1. Menghitung Ketidakpastian dengan banyak data 3
function DELTA_3(SIGMA, SIGMA_KUADRAT) {
  return hasil = (1 / 3) * (((3 * SIGMA_KUADRAT) - (SIGMA ** 2)) / (2)) ** 0.5;
}

function DELTA_5_COBA(SIGMA, SIGMA_KUADRAT) {
  return hasil = (1 / 5) * (((5 * SIGMA_KUADRAT - SIGMA ** 2) / (4)) ** 0.5);
}

// 2. Menghitung Ketidakpastian dengan banyak data 5
function DELTA_5(SIGMA, SIGMA_KUADRAT) {
  return hasil = (1 / 5) * (((5 * SIGMA_KUADRAT) - (SIGMA ** 2)) / (4)) ** 0.5;
}

// 3. Menghitung Ketidakpastian dengan banyak data n
function DELTA_N(n, SIGMA, SIGMA_KUADRAT) {
  return hasil = (1 / n) * (((n * SIGMA_KUADRAT) - (SIGMA ** 2)) / (n-1)) ** 0.5;
}

// 4. Menghitung Ketidakpastian (melewati beberapa step)
function DELTA_ADVANCED(DATA) {
  DATA = DATA.flat();
  var n = DATA.length; //dapet n

  let Sum_Sigma = 0;
  let Data_Kuadrat = 0;
  let Sum_Sigma_Kuadrat = 0;

  for (var i = 0; i < n; i++) {
    Sum_Sigma += DATA[i]; //dapet sigma n
    Data_Kuadrat = DATA[i] ** 2;
    Sum_Sigma_Kuadrat += Data_Kuadrat; //dapet sigma n^2
  }

  return hasil = (1 / n) * (((n * Sum_Sigma_Kuadrat) - (Sum_Sigma ** 2)) / (n-1)) ** 0.5;
}
