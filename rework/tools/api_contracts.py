"""Reviewed semantics of the pinned Swiss Ephemeris C interface.

ABI types still come from swephexp.h. These contracts describe caller storage,
result semantics and runnable examples; they never infer buffer size from types.
References: pinned source functions and https://www.astro.com/swisseph/swephprg.htm.
"""
from __future__ import annotations

# direction, allocation, visible result count. Over-allocation is deliberate for
# upstream reserved output slots; only documented initialized slots are exposed.
BUFFERS = {}
def buffers(names, **parameters):
    for name in names.split(): BUFFERS['swe_'+name] = parameters.copy()

def inp(n): return ('in',n,n)
def out(n,visible=None): return ('out',n,n if visible is None else visible)
def io(n): return ('inout',n,n)

buffers('azalt',geopos=inp(3),xin=inp(3),xaz=out(3))
buffers('azalt_rev',geopos=inp(3),xin=inp(2),xout=out(2))
buffers('calc calc_ut',xx=out(6))
buffers('calc_pctr',xxret=out(6))
buffers('cotrans',xpo=inp(3),xpn=out(3))
buffers('cotrans_sp',xpo=inp(6),xpn=out(6))
buffers('date_conversion',tjd=out(1))
buffers('fixstar fixstar_ut fixstar2 fixstar2_ut',xx=out(6))
buffers('fixstar_mag fixstar2_mag',mag=out(1))
buffers('gauquelin_sector',geopos=inp(3),dgsect=out(1))
buffers('get_ayanamsa_ex get_ayanamsa_ex_ut',daya=out(1))
buffers('get_current_file_data',tfstart=out(1),tfend=out(1),denum=out(1))
buffers('get_orbital_elements',dret=out(50,17))
buffers('heliacal_angle',dgeo=inp(3),datm=io(4),dobs=io(6),dret=out(3))
buffers('heliacal_pheno_ut',geopos=inp(3),datm=io(4),dobs=io(6),darr=out(50,28))
buffers('heliacal_ut',geopos=inp(3),datm=io(4),dobs=io(6),dret=out(10,3))
buffers('helio_cross helio_cross_ut',jd_cross=out(1))
buffers('house_pos',xpin=inp(2))
buffers('houses houses_armc houses_ex',cusps=out(37),ascmc=out(10,8))
buffers('houses_armc_ex2 houses_ex2',cusps=out(37),ascmc=out(10,8),cusp_speed=out(37),ascmc_speed=out(10,8))
for _house in ('swe_houses_armc','swe_houses_armc_ex2'):
    BUFFERS[_house]['ascmc']=('inout',10,10)
buffers('jdet_to_utc jdut1_to_utc',iyear=out(1),imonth=out(1),iday=out(1),ihour=out(1),imin=out(1),dsec=out(1))
buffers('lat_to_lmt',tjd_lmt=out(1))
buffers('lmt_to_lat',tjd_lat=out(1))
buffers('lun_eclipse_how sol_eclipse_how',geopos=inp(3),attr=out(20))
buffers('lun_eclipse_when sol_eclipse_when_glob lun_occult_when_glob',tret=out(10))
buffers('lun_eclipse_when_loc sol_eclipse_when_loc lun_occult_when_loc',geopos=inp(3),tret=out(10),attr=out(20))
buffers('lun_occult_where sol_eclipse_where',geopos=out(10,2),attr=out(20))
buffers('mooncross_node mooncross_node_ut',xlon=out(1),xlat=out(1))
buffers('nod_aps nod_aps_ut',xnasc=out(6),xndsc=out(6),xperi=out(6),xaphe=out(6))
buffers('orbit_max_min_true_distance',dmax=out(1),dmin=out(1),dtrue=out(1))
buffers('pheno pheno_ut',attr=out(20,6))
buffers('refrac_extended',dret=out(4))
buffers('revjul',jyear=out(1),jmon=out(1),jday=out(1),jut=out(1))
buffers('rise_trans rise_trans_true_hor',geopos=inp(3),tret=out(1))
buffers('split_deg',ideg=out(1),imin=out(1),isec=out(1),dsecfr=out(1),isgn=out(1))
buffers('time_equ',te=out(1))
buffers('topo_arcus_visionis',dgeo=inp(3),datm=io(4),dobs=io(6),dret=out(1))
buffers('utc_time_zone',iyear_out=out(1),imonth_out=out(1),iday_out=out(1),ihour_out=out(1),imin_out=out(1),dsec_out=out(1))
buffers('utc_to_jd',dret=out(2))
buffers('vis_limit_mag',geopos=inp(3),datm=io(4),dobs=io(6),dret=out(10,8))

# Each string buffer is writable memory with an explicit capacity, including NUL.
STRING_OUTPUTS={'swe_version':{'arg1'},'swe_get_library_path':{'arg1'},'swe_get_planet_name':{'spname'},'swe_cs2degstr':{'a'},'swe_cs2lonlatstr':{'s'},'swe_cs2timestr':{'a'},'swe_get_astro_models':{'sdet'}}
COMMANDS={'swe_close','swe_get_current_file_data'} | {'swe_set_'+s for s in ('astro_models delta_t_userdef ephe_path interpolate_nut jpl_file lapse_rate sid_mode tid_acc topo').split()}
FLAGS_RESULT=set('calc calc_ut calc_pctr fixstar fixstar_ut fixstar2 fixstar2_ut get_ayanamsa_ex get_ayanamsa_ex_ut'.split())
NO_EVENT=set('lun_eclipse_how sol_eclipse_how lun_eclipse_when sol_eclipse_when_glob lun_eclipse_when_loc sol_eclipse_when_loc lun_occult_when_glob lun_occult_when_loc lun_occult_where sol_eclipse_where'.split())
SIGNED_VALID=set('day_of_week csnorm csroundsec d2l difcs2n difcsn'.split())
CROSSINGS=set('solcross solcross_ut mooncross mooncross_ut mooncross_node mooncross_node_ut'.split())

UNITS={
 'geopos':'degrees east; degrees north; metres above sea level', 'dgeo':'degrees east; degrees north; metres above sea level',
 'datm':'hPa; Celsius; relative humidity percent; visibility km or extinction coefficient',
 'dobs':'age years; Snellen ratio; binocular flag; magnification; aperture mm; transmission fraction',
 'atpress':'hPa (0 estimates pressure)', 'attemp':'degrees Celsius', 'geoalt':'metres', 'horhgt':'degrees',
 'geolon':'degrees east', 'geolat':'degrees north', 'lapse_rate':'kelvin/metre', 'eps':'degrees', 'nut':'degrees', 'armc':'degrees',
 'ascmc':'house angles degrees; index 9 is Sun declination input for ARMC Sunshine houses','xpin':'ecliptic longitude and latitude, degrees', 'xpo':'longitude/latitude degrees; radius; optional matching daily speeds',
 'xin':'coordinates selected by calc_flag, degrees; optional radius', 'xpn':'longitude/latitude degrees; radius; optional matching daily speeds',
 'xaz':'azimuth from south toward west; true altitude; apparent altitude, degrees', 'xout':'longitude/right ascension; latitude/declination, degrees',
 'cusps':'degrees (radians with SEFLG_RADIANS); index 1..12 or 1..36 for G', 'ascmc':'degrees (radians with SEFLG_RADIANS)',
 'cusp_speed':'degrees/day (radians/day with SEFLG_RADIANS)', 'ascmc_speed':'degrees/day (radians/day with SEFLG_RADIANS)',
 'xx':'coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds',
 'xxret':'coordinate units selected by flags: degrees/radians or AU, followed by per-day speeds',
 'xnasc':'ascending node coordinates and speeds, selected by flags','xndsc':'descending node coordinates and speeds, selected by flags',
 'xperi':'periapsis coordinates and speeds, selected by flags','xaphe':'apoapsis/focal-point coordinates and speeds, selected by flags',
 't_ut':'Julian day UT1','tjdut':'Julian day UT1','mag':'astronomical magnitude','daya':'degrees','dgsect':'Gauquelin sector 1..36','tret':'Julian day UT1; 0 means unavailable contact',
 'xlon':'degrees','xlat':'degrees','dmax':'AU','dmin':'AU','dtrue':'AU','te':'days (apparent minus mean solar time)',
 'dt':'days; automatic sentinel -1e-10','t_acc':'arcseconds/century squared; automatic sentinel 999999',
 'tfstart':'Julian day TT','tfend':'Julian day TT','denum':'JPL DE number', 'd_timezone':'hours east of UTC (fractional allowed)',
 'dsec':'seconds, leap second allowed where documented','dsec_out':'seconds','dsecfr':'fractional arcsecond',
 'year':'astronomical year','y':'astronomical year','iyear':'astronomical year','jyear':'astronomical year','iyear_out':'astronomical year',
 'month':'month 1..12','m':'month 1..12','imonth':'month 1..12','jmon':'month 1..12','imonth_out':'month 1..12',
 'day':'day of month','d':'day of month','iday':'day of month','jday':'day of month','iday_out':'day of month',
 'hour':'decimal hours','utime':'decimal hours','jut':'decimal hours','ihour':'hour 0..23','ihour_out':'hour 0..23',
 'imin':'minutes','imin_out':'minutes','ideg':'whole degrees','isec':'arcseconds','isgn':'sign or zodiac index selected by roundflag',
 'hsys':'one-character house system','c':'g Gregorian or j Julian','gregflag':'0 Julian, 1 Gregorian',
 'ipl':'Swiss body ID','iplctr':'Swiss centre body ID','ifno':'file slot 0..4','isidmode':'built-in sidereal ID 0..46',
 'sid_mode':'sidereal mode ID and modifier bits; 255 user defined','t0':'reference Julian day TT unless sidereal modifier changes it',
 'ayan_t0':'degrees','iflag':'Swiss flag bit mask','ifl':'ephemeris flag bit mask','epheflag':'ephemeris flag bit mask','flag':'Swiss flag bit mask',
 'helflag':'heliacal flag bit mask','rsmi':'rise/set/transit bit mask','ifltype':'event type bit mask','method':'node/apsis method bit mask',
 'imeth':'Gauquelin method 0..5','TypeEvent':'heliacal event type 1..6','roundflag':'split-degree bit mask','backward':'0 forward, 1 backward; occultation one-try bit 32768',
 'dir':'positive forward or negative backward','do_interpolate':'C boolean 0 or 1','suppressZero':'C boolean 0 or 1',
 'calc_flag':'coordinate/refraction mode code','x2cross':'target ecliptic longitude, degrees', 'inalt':'altitude degrees','ddeg':'degrees',
 'azi_obj':'azimuth degrees','azi_sun':'azimuth degrees','azi_moon':'azimuth degrees','alt_obj':'altitude degrees','alt_moon':'altitude degrees',
 'path':'absolute directory or path list, Windows ANSI encoding','fname':'JPL filename relative to configured data path',
 'samod':'experimental astronomical model/version specification','sdet':'experimental model-description text',
 'star':'star name or catalog index; canonical name is returned','starname':'star name, or empty to use ipl','ObjectName':'planet or fixed-star name',
 'spname':'body name','arg1':'NUL-terminated string','a':'formatted text','s':'formatted text','pchar':'positive hemisphere character','mchar':'negative hemisphere character','sep':'separator character code',
}

ARRAY_UNITS={
 ('utc_to_jd','dret'):'Julian day TT, Julian day UT1',
 ('get_orbital_elements','dret'):'a AU; e; inclination/node/periapsis/anomalies/longitudes degrees; daily motion degrees/day; sidereal/tropical periods years; synodic period days; perihelion epoch JD TT; perihelion/aphelion AU',
 ('refrac_extended','dret'):'true altitude deg; apparent altitude deg; refraction deg; dip deg',
 ('heliacal_ut','dret'):'Julian day UT1 beginning/optimum/end; unavailable dates are zero',
 ('heliacal_angle','dret'):'optimum object altitude; minimum arcus visionis; solar altitude, degrees',
 ('topo_arcus_visionis','dret'):'arcus visionis degrees',
 ('vis_limit_mag','dret'):'limiting magnitude; object alt/az deg; Sun alt/az deg; Moon alt/az deg; object magnitude',
 ('heliacal_pheno_ut','darr'):'indexed heliacal circumstances: degrees, days, magnitudes and dimensionless visibility measures; see field reference',
}

DEFAULTS={'tjd':2451545.0,'tjd_ut':2451545.0,'tjd_et':2451545.0,'t_ut':2451545.0,'tjdut':2451545.0,'tjd_start':2451545.0,'tjdstart_ut':2451545.0,'jd_et':2451545.0,'jd_ut':2451545.0,'jd':2451545.0,'tjd_lat':2451545.0,'tjd_lmt':2451545.0,
 'ipl':0,'iplctr':4,'iflag':258,'ifl':2,'epheflag':2,'flag':258,'helflag':0,'ifltype':0,'backward':0,'dir':1,'gregflag':1,
 'year':2000,'y':2000,'iyear':2000,'month':1,'m':1,'imonth':1,'day':1,'d':1,'iday':1,'hour':12.0,'utime':12.0,'ihour':12,'imin':0,'dsec':0.0,'d_timezone':5.5,
 'geolon':13.405,'geolat':52.52,'geoalt':34.0,'atpress':1013.25,'attemp':15.0,'lapse_rate':0.0065,'horhgt':0.0,
 'ascmc':[0.0]*10,'geopos':[13.405,52.52,34.0],'dgeo':[13.405,52.52,34.0],'datm':[1013.25,15.0,40.0,0.0],'dobs':[36.0,1.0,0.0,1.0,0.0,0.0],
 'xpo':[280.0,1.0,1.0],'xin':[280.0,1.0,1.0],'xpin':[280.0,1.0],'eps':23.4392911,'nut':0.0,'armc':280.0,'hsys':'P','c':'g',
 'star':'Spica','starname':'','ObjectName':'venus','samod':'0,0,0,0,0,0,0,0','spname':'','pchar':'N','mchar':'S','sep':58,
 'x':12.345,'x1':350.0,'x0':10.0,'p':-360000,'p1':129600000,'p2':360000,'t':1234567,'inalt':5.0,'ddeg':123.456,
 'suppressZero':0,'isidmode':1,'sid_mode':1,'t0':2451545.0,'ayan_t0':24.0,'dt':-1e-10,'t_acc':999999.0,'do_interpolate':0,'ifno':0,
 'imeth':0,'method':0,'TypeEvent':1,'rsmi':1,'roundflag':0,'calc_flag':0,'x2cross':0.0,'mag':0.0,
 'azi_obj':90.0,'azi_sun':100.0,'azi_moon':110.0,'alt_obj':5.0,'alt_moon':-10.0,'path':'@DATA@','fname':'de431.eph'}

def describe(function):
    name=function['name']; short=name[4:]; params=[]
    for p in function['parameters']:
        arg=p['name'];char=p['cType'].replace('const ','')=='char *'
        if char:
            direction='out' if arg=='serr' or arg in STRING_OUTPUTS.get(name,set()) else 'inout' if arg=='star' else 'in'
            size=16384 if arg=='sdet' else 256
            d=(direction,size,1)
        elif p['pointer']:
            if arg not in BUFFERS.get(name,{}):raise ValueError(f'Unreviewed pointer {name}.{arg}')
            d=BUFFERS[name][arg]
        else:d=('in',1,1)
        unit=UNITS.get(arg)
        if p['cType'] in ('centisec','CSEC'):unit='centiseconds (1/360000 degree)'
        elif arg in ('x','x1','x0','p1','p2'):unit='radians' if 'rad' in short else 'degrees'
        elif arg.startswith(('tjd','jd_')) or arg=='jd':unit='Julian day UT1' if '_ut' in arg or arg in ('t_ut','tjdut','tjd_start','tjdstart_ut') else 'Julian day TT or explicitly named input time scale'
        elif arg=='serr':unit='diagnostic text'
        if (short,arg) in ARRAY_UNITS:unit=ARRAY_UNITS[short,arg]
        if arg=='attr':unit='indexed event/phenomena attributes; see field reference'
        if unit is None:raise ValueError(f'Unreviewed unit {name}.{arg}')
        example=DEFAULTS.get(arg)
        if d[0]!='out' and example is None:raise ValueError(f'Missing example {name}.{arg}')
        if arg in ('p1','p2') and p['cType'] not in ('centisec','CSEC'):example=({'p1':3.0,'p2':-3.0} if 'rad' in short else {'p1':350.0,'p2':10.0})[arg]
        if arg=='tjd' and (any(x in short for x in ('eclipse','occult')) or short in ('time_equ','sidtime','sidtime0','azalt','azalt_rev')):unit='Julian day UT1'
        if arg=='jd_cross' and short.endswith('_ut'):unit='Julian day UT1'
        if short=='csroundsec' and arg=='x':example=1234567
        if short=='cotrans_sp' and arg=='xpo':example=[280.0,1.0,1.0,1.0,0.0,0.0]
        if short=='azalt_rev' and arg=='xin':example=[90.0,15.0]
        if short=='get_orbital_elements' and arg=='ipl':example=4
        if short in ('helio_cross','helio_cross_ut') and arg=='ipl':example=4
        if short.startswith('lun_occult') and arg=='ipl':example=3
        params.append(dict(p,direction=d[0],capacity=d[1],visible=d[2],units=unit,string=char,example=example))
    kind='flags' if short in FLAGS_RESULT else 'event' if short in NO_EVENT else 'crossing' if short in CROSSINGS else 'signed' if short in SIGNED_VALID else 'status' if function['returns']['vbaType']=='Long' else 'value'
    if short=='house_pos':kind='house_position'
    if short=='vis_limit_mag':kind='visibility'
    if short.startswith('rise_trans'):kind='rise'
    family=('Configuration' if name in COMMANDS else 'Houses' if 'house' in short or short=='gauquelin_sector' else 'Events' if any(x in short for x in ('eclipse','occult','rise_trans','cross')) else 'Heliacal' if any(x in short for x in ('heliacal','vis_limit','arcus')) else 'Stars' if 'fixstar' in short else 'Time' if any(x in short for x in ('utc','jul','date','deltat','time','lmt','lat_to','day_of_week','tid_acc')) else 'Positions' if any(x in short for x in ('calc','pheno','nod_aps','orbit','ayanamsa','planet_name')) else 'Utilities')
    required = ('No binary data normally required; delta-T/leap-second model applies to time conversions.' if family in ('Utilities','Time') else 'Body/date/flags determine files: sepl_18.se1, semo_18.se1, seas_18.se1; external asteroid/JPL files where requested.')
    if family=='Stars':required='sefstars.txt; selected planetary ephemeris for apparent corrections.'
    if family=='Configuration':required='Configuration/state query; supplied paths/files must exist for subsequent calculations.'
    if family=='Houses':required='Date-based houses use obliquity/nutation; ARMC houses need no binary data; Gauquelin body calculations need the selected body files.'
    return {'requiredFiles':required,'name':name,'worksheetName':('SW_CMD_' if name in COMMANDS else 'SW_SWE_')+short.upper(),'command':name in COMMANDS,'family':family,'parameters':params,'resultKind':kind,'reference':'vendor/swisseph/source: '+name+'; https://www.astro.com/swisseph/swephprg.htm'}
