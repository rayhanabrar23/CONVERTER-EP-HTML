// Pengaturan versi aplikasi — satu-satunya file yang berbeda antara versi EP dan versi umum.
window.PEI_CONFIG = {
  variant: 'ep',
  title: 'PEI Generator — EP · MNC Sekuritas',
  chip: '🏦 Partisipan EP · MNC Sekuritas',
  participantCode: 'EP',        // dikunci; berkas dengan Broker Code lain ditolak
  participantLabel: 'EP (MNC Sekuritas)',
  clPartisipan: 160000000000,   // Credit Limit partisipan di PEI (default, bisa diubah di Langkah 1)
  partnerLabel: 'MNC',          // sebutan pihak partisipan di pesan
  portoRate: 9.5,               // bunga (% p.a.) untuk transaksi baru di porto; transaksi lama tetap memakai tarif saat itu
};
