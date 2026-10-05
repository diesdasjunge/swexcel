"""Indexed output meanings reviewed against pinned swecl/swehel/swehouse/sweph C."""
FIELDS={}
def put(names,arg,labels):
 for name in names.split():FIELDS['swe_'+name,arg]=labels.split('|')
put('utc_to_jd','dret','Julian day TT|Julian day UT1')
put('get_orbital_elements','dret','semimajor axis AU|eccentricity|inclination degrees|ascending node degrees|argument of perihelion degrees|longitude of perihelion degrees|mean anomaly degrees|true anomaly degrees|eccentric anomaly degrees|mean longitude degrees|sidereal period tropical years|mean daily motion degrees/day|tropical period years|synodic period days|perihelion epoch JD TT|perihelion distance AU|aphelion distance AU')
put('pheno pheno_ut','attr','phase angle degrees|illuminated fraction|elongation degrees|apparent diameter degrees|apparent magnitude|horizontal parallax degrees (Moon)')
put('refrac_extended','dret','true altitude degrees|apparent altitude degrees|refraction degrees|horizon dip degrees')
put('heliacal_ut','dret','first visibility JD UT1|optimum visibility JD UT1|last visibility JD UT1')
put('heliacal_angle','dret','optimum object altitude degrees|minimum arcus visionis degrees|Sun altitude degrees')
put('vis_limit_mag','dret','limiting magnitude|object altitude degrees|object azimuth degrees|Sun altitude degrees|Sun azimuth degrees|Moon altitude degrees|Moon azimuth degrees|object magnitude')
put('heliacal_pheno_ut','darr','object true altitude degrees|object apparent altitude degrees|geocentric altitude degrees|object azimuth degrees|Sun altitude degrees|Sun azimuth degrees|topocentric arcus visionis degrees|geocentric arcus visionis degrees|azimuth difference degrees|longitude difference degrees|extinction coefficient|minimum arcus visionis degrees|first visibility JD UT1|optimum visibility JD UT1|last visibility JD UT1|Yallop best time JD UT1|crescent width degrees|Yallop q|Yallop criterion|parallax degrees|magnitude|object rise/set JD UT1|Sun rise/set JD UT1|lag days|visibility duration days|crescent length degrees|elongation degrees|illuminated percent')
put('azalt','xaz','azimuth south through west degrees|true altitude degrees|apparent altitude degrees')
put('sol_eclipse_where lun_occult_where','geopos','central longitude degrees east|central latitude degrees north')
solar='diameter fraction covered|Moon/object diameter ratio|disc obscuration fraction|core shadow diameter km|azimuth south through west degrees|true altitude degrees|apparent altitude degrees|Moon/object separation degrees|NASA magnitude|Saros series (-99999999 unavailable)|Saros member (-99999999 unavailable)'
put('sol_eclipse_where sol_eclipse_how sol_eclipse_when_loc lun_occult_where lun_occult_when_loc','attr',solar)
put('lun_eclipse_how lun_eclipse_when_loc','attr','umbral magnitude|penumbral magnitude|reserved|reserved|Moon azimuth degrees|true altitude degrees|apparent altitude degrees|opposition distance degrees|eclipse magnitude|Saros series|Saros member')
put('sol_eclipse_when_glob lun_occult_when_glob','tret','maximum JD UT1|local apparent noon JD UT1|eclipse begins JD UT1|eclipse ends JD UT1|totality begins JD UT1|totality ends JD UT1|centre line begins JD UT1|centre line ends JD UT1|reserved unimplemented|reserved unimplemented')
put('sol_eclipse_when_loc lun_occult_when_loc','tret','maximum JD UT1|first contact JD UT1|second contact JD UT1|third contact JD UT1|fourth contact JD UT1|sunrise JD UT1|sunset JD UT1|reserved|reserved|reserved')
put('lun_eclipse_when lun_eclipse_when_loc','tret','maximum JD UT1|reserved|partial begins JD UT1|partial ends JD UT1|total begins JD UT1|total ends JD UT1|penumbral begins JD UT1|penumbral ends JD UT1|moonrise JD UT1 (local only)|moonset JD UT1 (local only)')
ANGLES='Ascendant|Midheaven|ARMC|Vertex|equatorial Ascendant|Koch co-Ascendant|Munkasey co-Ascendant|polar Ascendant|reserved|Sun declination (Sunshine ARMC input)'.split('|')
def fields(name,p):
 if (name,p['name']) in FIELDS:return FIELDS[name,p['name']]
 if p['name'] in ('ascmc','ascmc_speed'):return [x+(' daily speed' if p['name'].endswith('speed') else '')+'; '+p['units'] for x in ANGLES]
 if p['name']=='attr':return ['reserved / not implemented']*p['visible']
 return None
