-- PP2 BOSQUE NE — Edificios Extra v2
-- Genera: Puesto Científico + Torre de Radio + Campamento Militar
-- Requiere: BosqueNE_Zone en Workspace | Suelo Y=331

local ws = game:GetService("Workspace")
local zona = ws:FindFirstChild("BosqueNE_Zone")
assert(zona, "Crea la carpeta BosqueNE_Zone en Workspace primero")

for _, n in ipairs({"PP2_PuestoCientifico","PP2_TorreRadio","PP2_CampamentoMilitar"}) do
	local old = zona:FindFirstChild(n)
	if old then old:Destroy() end
end

local GY = 331

-- ── helpers ─────────────────────────────────────────────────────────────────

local function pt(par,nm, sx,sy,sz, x,y,z, col, mat, tr)
	local p = Instance.new("Part")
	p.Name=nm; p.Size=Vector3.new(sx,sy,sz)
	p.CFrame=CFrame.new(x,y,z)
	p.BrickColor=BrickColor.new(col)
	p.Material=mat or Enum.Material.SmoothPlastic
	p.Transparency=tr or 0; p.Anchored=true; p.Parent=par
	return p
end

local function wd(par,nm, sx,sy,sz, cf, col, mat)
	local p = Instance.new("WedgePart")
	p.Name=nm; p.Size=Vector3.new(sx,sy,sz); p.CFrame=cf
	p.BrickColor=BrickColor.new(col)
	p.Material=mat or Enum.Material.SmoothPlastic
	p.Anchored=true; p.Parent=par; return p
end

local function md(par,nm)
	local m=Instance.new("Model"); m.Name=nm; m.Parent=par; return m
end

local function gl(p, br,rng, r,g,b)
	local l=Instance.new("PointLight")
	l.Brightness=br; l.Range=rng; l.Color=Color3.fromRGB(r,g,b); l.Parent=p
end

local function pp(p, txt)
	local pr=Instance.new("ProximityPrompt")
	pr.ActionText=txt; pr.MaxActivationDistance=8; pr.Parent=p
end

local function smk(p, sz, op)
	local s=Instance.new("Smoke")
	s.Size=sz; s.Opacity=op; s.RiseVelocity=5
	s.Color=Color3.fromRGB(70,70,70); s.Parent=p
end

local function fir(p, sz, ht)
	local f=Instance.new("Fire"); f.Size=sz; f.Heat=ht; f.Parent=p
end

local function sgn(p, txt, r,g,b)
	local sg=Instance.new("SurfaceGui",p); sg.Face=Enum.NormalId.Front
	local lb=Instance.new("TextLabel",sg)
	lb.Size=UDim2.new(1,0,1,0); lb.BackgroundTransparency=1
	lb.Text=txt; lb.TextScaled=true; lb.Font=Enum.Font.GothamBold
	lb.TextColor3=Color3.fromRGB(r or 255, g or 220, b or 120)
end

local function itm(par,nm, x,y,z, txt)
	local i=pt(par,nm, 3,3,3, x,y,z, "Bright yellow",Enum.Material.Neon)
	gl(i,4,14,255,230,60); pp(i,txt); return i
end

-- ════════════════════════════════════════════════════════════════════════════
--  1. PUESTO CIENTÍFICO  (500, GY, -420)
-- ════════════════════════════════════════════════════════════════════════════
local sci = md(zona,"PP2_PuestoCientifico")
local SX,SZ = 500,-420

pt(sci,"Suelo", 160,2,130, SX,GY-1,SZ, "Moss",Enum.Material.Grass)

-- Laboratorio (70×70 base, 20 alto)
local lab = md(sci,"Laboratorio")
pt(lab,"N",  70,20,3, SX,    GY+10, SZ-35, "Sand green",Enum.Material.Concrete)
pt(lab,"S",  70,20,3, SX,    GY+10, SZ+35, "Sand green",Enum.Material.Concrete)
pt(lab,"E",  3, 20,70, SX+35, GY+10, SZ,   "Sand green",Enum.Material.Concrete)
pt(lab,"O",  3, 20,70, SX-35, GY+10, SZ,   "Sand green",Enum.Material.Concrete)
pt(lab,"Piso",70,2,70, SX,    GY+1,  SZ,   "Light grey",Enum.Material.SmoothPlastic)
pt(lab,"Techo",72,3,72,SX,    GY+21, SZ,   "Sand green",Enum.Material.Metal)
-- Ventanas
for _,ox in ipairs({-22,-8,8,22}) do
	pt(lab,"VN"..ox, 10,8,0.4, SX+ox,GY+13,SZ-35.2, "Medium stone grey",Enum.Material.Glass,0.3)
	pt(lab,"VS"..ox, 10,8,0.4, SX+ox,GY+13,SZ+35.2, "Medium stone grey",Enum.Material.Glass,0.3)
end
-- Puerta
pt(lab,"Puerta",10,14,1, SX,GY+8,SZ+35.5, "Dark grey",Enum.Material.Metal)
-- Mesas de laboratorio con microscopios
for i=1,4 do
	local mx = SX-22+(i*12)
	pt(lab,"Mesa"..i,  14,4,6,  mx,GY+3, SZ-16, "Light grey",Enum.Material.SmoothPlastic)
	pt(lab,"Micro"..i,  3,8,3,  mx,GY+7, SZ-16, "Dark grey", Enum.Material.Metal)
	pt(lab,"MTop"..i,   5,2,5,  mx,GY+11,SZ-16, "Dark grey", Enum.Material.Metal)
	-- Tubo de ensayo con luz verde
	local tb = pt(lab,"Tubo"..i, 2,6,2, mx,GY+7,SZ+14, "Bright green",Enum.Material.Neon,0.3)
	pt(lab,"Mesa2_"..i,14,4,6,  mx,GY+3, SZ+14, "Light grey",Enum.Material.SmoothPlastic)
	gl(tb,1,6,60,255,80)
end
-- Estantes laterales
pt(lab,"EstE", 3,16,60, SX+33,GY+9,SZ, "Dark grey",Enum.Material.Metal)
pt(lab,"EstO", 3,16,60, SX-33,GY+9,SZ, "Dark grey",Enum.Material.Metal)
-- Pizarrón
local piz = pt(lab,"Pizarron",30,12,1, SX,GY+12,SZ-34, "Dark grey",Enum.Material.SmoothPlastic)
sgn(piz,"PROYECTO ECO — DATOS CLASIFICADOS",60,255,120)
-- Item PP2 #6
itm(lab,"PP2_Item6", SX-22,GY+8,SZ-16, "Recolectar Muestra")

-- Generador roto exterior
local gen = md(sci,"Generador")
pt(gen,"Cuerpo",14,12,10, SX-60,GY+7,SZ-30,"Dark grey",Enum.Material.Metal)
pt(gen,"Tapa",  16,3,12,  SX-60,GY+14,SZ-30,"Very dark grey",Enum.Material.Metal)
local gtop = pt(gen,"HumoGen",6,2,6, SX-60,GY+16,SZ-30,"Very dark grey",Enum.Material.Metal)
smk(gtop,1.5,0.8)
pt(gen,"Tuberia",3,3,22, SX-60,GY+8,SZ-19,"Dark grey",Enum.Material.Metal)
local gf = pt(gen,"Chispa",3,3,3, SX-54,GY+8,SZ-30,"Bright yellow",Enum.Material.Neon)
fir(gf,3,3); gl(gf,3,12,255,200,50)

-- Antena en techo
pt(sci,"Mastil",   2,30,2, SX+20,GY+37,SZ-20,"Dark grey",Enum.Material.Metal)
pt(sci,"BrazoH1",  18,2,2, SX+20,GY+54,SZ-20,"Dark grey",Enum.Material.Metal)
pt(sci,"BrazoH2",  10,2,2, SX+20,GY+48,SZ-20,"Dark grey",Enum.Material.Metal)
pt(sci,"BrazoV",   2,2,18, SX+20,GY+54,SZ-20,"Dark grey",Enum.Material.Metal)
local lr = pt(sci,"LuzAntena",2,2,2, SX+20,GY+56,SZ-20,"Bright red",Enum.Material.Neon)
gl(lr,3,20,255,40,40)

-- Jaulas de animales
for i=1,3 do
	local jx = SX+42+(i*18)
	pt(sci,"JPiso"..i, 14,2,14,  jx,GY+1, SZ+40,"Dark grey",Enum.Material.Metal)
	pt(sci,"JL"..i,    2,14,14,  jx-7,GY+8,SZ+40,"Dark grey",Enum.Material.Metal)
	pt(sci,"JR"..i,    2,14,14,  jx+7,GY+8,SZ+40,"Dark grey",Enum.Material.Metal)
	pt(sci,"JN"..i,    14,14,2,  jx,GY+8, SZ+33,"Dark grey",Enum.Material.Metal)
	pt(sci,"JS"..i,    14,14,2,  jx,GY+8, SZ+47,"Dark grey",Enum.Material.Metal)
	pt(sci,"JT"..i,    14,2,14,  jx,GY+15,SZ+40,"Dark grey",Enum.Material.Metal)
end

-- Letrero entrada
pt(sci,"PosteL",3,16,3, SX-12,GY+8,SZ+65,"Dark grey",Enum.Material.Metal)
pt(sci,"PosteR",3,16,3, SX+12,GY+8,SZ+65,"Dark grey",Enum.Material.Metal)
local ls = pt(sci,"Letrero",30,7,2, SX,GY+18,SZ+65,"Dark grey",Enum.Material.Metal)
sgn(ls,"ESTACION BIOLOGICA EKO-7",60,255,140)

print("Puesto Cientifico generado en ("..SX..","..GY..","..SZ..")")

-- ════════════════════════════════════════════════════════════════════════════
--  2. TORRE DE RADIO  (780, GY, -580)
-- ════════════════════════════════════════════════════════════════════════════
local rad = md(zona,"PP2_TorreRadio")
local RX,RZ = 780,-580

pt(rad,"Suelo",120,2,120, RX,GY-1,RZ,"Dark grey",Enum.Material.Concrete)

-- Sala de control (44×44 base, 18 alto)
local ctl = md(rad,"SalaControl")
pt(ctl,"N",  44,18,3, RX,    GY+9,RZ-25,"Very dark grey",Enum.Material.Metal)
pt(ctl,"S",  44,18,3, RX,    GY+9,RZ+25,"Very dark grey",Enum.Material.Metal)
pt(ctl,"E",  3, 18,50, RX+22, GY+9,RZ,  "Very dark grey",Enum.Material.Metal)
pt(ctl,"O",  3, 18,50, RX-22, GY+9,RZ,  "Very dark grey",Enum.Material.Metal)
pt(ctl,"Piso",44,2,50, RX,    GY+1,RZ,  "Dark grey",Enum.Material.Concrete)
pt(ctl,"Techo",46,3,52,RX,    GY+19,RZ, "Very dark grey",Enum.Material.Metal)
-- Ventana grande
pt(ctl,"Ventana",30,10,0.4, RX,GY+12,RZ+25.2,"Medium stone grey",Enum.Material.Glass,0.2)
-- Puerta
pt(ctl,"Puerta",10,14,1, RX-14,GY+8,RZ-25.5,"Dark grey",Enum.Material.Metal)
-- Consolas con pantallas
for i=1,5 do
	local cx = RX-16+(i*7)
	pt(ctl,"Consola"..i, 8,6,5,  cx,GY+4,RZ-16,"Very dark grey",Enum.Material.Metal)
	local scr = pt(ctl,"Pantalla"..i, 6,4,0.3, cx,GY+5,RZ-16.2,"Bright blue",Enum.Material.Neon,0.1)
	gl(scr,1,8,40,120,255)
end
pt(ctl,"Silla",5,5,5, RX,GY+3,RZ,"Dark grey",Enum.Material.SmoothPlastic)
local mapa = pt(ctl,"MapaPared",34,12,0.4, RX,GY+12,RZ+24.8,"Sand green",Enum.Material.SmoothPlastic)
sgn(mapa,"MAPA DE COBERTURA - CLASIFICADO",60,255,140)
itm(ctl,"PP2_Item7", RX,GY+7,RZ,"Recolectar Codigos de Radio")

-- Torre antena principal (100 studs)
local TAX,TAZ = RX-45,RZ-40
pt(rad,"TorreMastil",4,100,4, TAX,GY+50,TAZ,"Dark grey",Enum.Material.Metal)
-- Patas diagonales
for _,ox in ipairs({-12,12}) do
	for _,oz in ipairs({-12,12}) do
		wd(rad,"Pata"..ox..oz, 3,40,3,
			CFrame.new(TAX+ox*0.5, GY+20, TAZ+oz*0.5) * CFrame.Angles(0,0,math.rad(ox>0 and -10 or 10)),
			"Dark grey",Enum.Material.Metal)
	end
end
-- Plataformas intermedias con barandal
for _,h in ipairs({28,54,80}) do
	pt(rad,"TPlat"..h,  20,2,20, TAX,GY+h,TAZ,"Dark grey",Enum.Material.Metal)
	pt(rad,"TRN"..h,    20,3,2,  TAX,GY+h+2,TAZ-10,"Dark grey",Enum.Material.Metal)
	pt(rad,"TRS"..h,    20,3,2,  TAX,GY+h+2,TAZ+10,"Dark grey",Enum.Material.Metal)
	pt(rad,"TRE"..h,    2,3,20,  TAX+10,GY+h+2,TAZ,"Dark grey",Enum.Material.Metal)
	pt(rad,"TRO"..h,    2,3,20,  TAX-10,GY+h+2,TAZ,"Dark grey",Enum.Material.Metal)
end
-- Antenas en cima
pt(rad,"AntV",  2,20,2, TAX,    GY+111,TAZ,    "Dark grey",Enum.Material.Metal)
pt(rad,"AntH1", 30,2,2, TAX,    GY+106,TAZ,    "Dark grey",Enum.Material.Metal)
pt(rad,"AntH2", 2,2,30, TAX,    GY+106,TAZ,    "Dark grey",Enum.Material.Metal)
pt(rad,"AntH3", 18,2,2, TAX,    GY+100,TAZ,    "Dark grey",Enum.Material.Metal)
local lrcima = pt(rad,"LuzCima",3,3,3, TAX,GY+121,TAZ,"Bright red",Enum.Material.Neon)
gl(lrcima,5,35,255,30,30)
-- Cables
pt(rad,"Cable1",55,1,1, RX-10,GY+22,RZ-10,"Very dark grey",Enum.Material.Metal)
pt(rad,"Cable2",55,1,1, RX-10,GY+22,RZ+10,"Very dark grey",Enum.Material.Metal)
-- Generador
pt(rad,"GenCuerpo",16,10,10, RX+35,GY+6,RZ-35,"Dark grey",Enum.Material.Metal)
pt(rad,"GenTapa",  18,3,12,  RX+35,GY+12,RZ-35,"Very dark grey",Enum.Material.Metal)
local gh = pt(rad,"GenHumo",5,2,5, RX+35,GY+14,RZ-35,"Very dark grey",Enum.Material.Metal)
smk(gh,1,0.7)

pt(rad,"PosteL",3,16,3, RX-12,GY+8,RZ+55,"Dark grey",Enum.Material.Metal)
pt(rad,"PosteR",3,16,3, RX+12,GY+8,RZ+55,"Dark grey",Enum.Material.Metal)
local rls = pt(rad,"Letrero",32,7,2, RX,GY+18,RZ+55,"Dark grey",Enum.Material.Metal)
sgn(rls,"ESTACION DE RADIO EKO",255,220,60)

print("Torre de Radio generada en ("..RX..","..GY..","..RZ..")")

-- ════════════════════════════════════════════════════════════════════════════
--  3. CAMPAMENTO MILITAR  (900, GY, -200)
-- ════════════════════════════════════════════════════════════════════════════
local mil = md(zona,"PP2_CampamentoMilitar")
local MX,MZ = 900,-200

pt(mil,"Suelo", 200,2,180, MX,GY-1,MZ,"Sand green",Enum.Material.Grass)
pt(mil,"Camino",  8,2,90,  MX,GY,  MZ,"Dark grey",Enum.Material.Concrete)

-- Búnker central (60×50 base, semi-enterrado)
local bun = md(mil,"Bunker")
pt(bun,"N",  60,16,4, MX,    GY+8,MZ-28,"Dark grey",Enum.Material.Concrete)
pt(bun,"S",  60,16,4, MX,    GY+8,MZ+28,"Dark grey",Enum.Material.Concrete)
pt(bun,"E",  4, 16,56, MX+30, GY+8,MZ,  "Dark grey",Enum.Material.Concrete)
pt(bun,"O",  4, 16,56, MX-30, GY+8,MZ,  "Dark grey",Enum.Material.Concrete)
pt(bun,"Piso",60,2,56, MX,    GY+1,MZ,  "Dark grey",Enum.Material.Concrete)
pt(bun,"Techo",64,6,60,MX,    GY+17,MZ, "Dark grey",Enum.Material.Concrete)
pt(bun,"Tierra",62,4,58,MX,   GY+21,MZ, "Moss",Enum.Material.Grass)
-- Puerta
pt(bun,"Puerta",12,12,1, MX,GY+7,MZ+28.5,"Very dark grey",Enum.Material.Metal)
-- Interior
pt(bun,"MesaTactica",22,4,12, MX,GY+3,MZ-8,"Dark grey",Enum.Material.Metal)
local mapaT = pt(bun,"MapaTactico",20,12,0.4, MX,GY+10,MZ-27.8,"Sand green",Enum.Material.SmoothPlastic)
sgn(mapaT,"MAPA TACTICO - OPERACION ECO",255,220,60)
pt(bun,"Radio",6,5,5, MX+18,GY+4,MZ-8,"Dark grey",Enum.Material.Metal)
for i=1,4 do
	pt(bun,"Litera"..i, 10,4,20, MX-10+(i*7),GY+3,MZ+14,"Dark grey",Enum.Material.Metal)
end
local lem = pt(bun,"LuzEmergencia",3,3,3, MX,GY+15,MZ,"Bright red",Enum.Material.Neon)
gl(lem,2,18,255,30,30)
itm(bun,"PP2_Item8", MX,GY+7,MZ-8,"Recolectar Planos Militares")

-- Tiendas de campaña (3)
for i=1,3 do
	local tz = MZ-50+(i*36)
	local tie = md(mil,"Tienda"..i)
	pt(tie,"Base", 22,1,16, MX-70,GY+1,tz,"Dark olive green",Enum.Material.SmoothPlastic)
	wd(tie,"TL",   22,10,9, CFrame.new(MX-70,GY+6,tz-4),"Dark olive green",Enum.Material.SmoothPlastic)
	wd(tie,"TR",   22,10,9, CFrame.new(MX-70,GY+6,tz+4)*CFrame.Angles(0,math.pi,0),"Dark olive green",Enum.Material.SmoothPlastic)
	pt(tie,"Poste",2,12,2, MX-70+12,GY+7,tz,"Reddish brown",Enum.Material.Wood)
end

-- Jeep oxidado
local jeep = md(mil,"JeepOxidado")
local JX,JZ = MX+65,MZ+35
pt(jeep,"Cuerpo",  28,10,14, JX,    GY+6, JZ,  "Olive",Enum.Material.Metal)
pt(jeep,"Capo",    16,4,14,  JX-16, GY+9, JZ,  "Olive",Enum.Material.Metal)
pt(jeep,"Cabina",  14,8,12,  JX+4,  GY+12,JZ,  "Olive",Enum.Material.Metal)
for _,ox in ipairs({-10,10}) do
	for _,oz in ipairs({-8,8}) do
		pt(jeep,"Rueda"..ox..oz,8,8,4, JX+ox,GY+4,JZ+oz,"Very dark grey",Enum.Material.SmoothPlastic)
	end
end
pt(jeep,"Parabrisas",14,6,0.4, JX+4,GY+14,JZ-6,"Medium stone grey",Enum.Material.Glass,0.4)

-- Rampa entrada búnker
wd(mil,"Rampa",14,5,14,
	CFrame.new(MX,GY+2,MZ+36)*CFrame.Angles(math.rad(-20),0,0),
	"Dark grey",Enum.Material.Concrete)

-- Barricadas + alambre
for i=1,6 do
	pt(mil,"Barricada"..i,10,8,4, MX-75+(i*14),GY+5,MZ+85,"Dark grey",Enum.Material.Concrete)
end
pt(mil,"Alambre1",90,1,1, MX-25,GY+10,MZ+85,"Dark grey",Enum.Material.Metal)
pt(mil,"Alambre2",90,1,1, MX-25,GY+7, MZ+85,"Dark grey",Enum.Material.Metal)

-- Focos exteriores en postes
for _,pos in ipairs({{MX-90,MZ-80},{MX+90,MZ-80},{MX-90,MZ+80},{MX+90,MZ+80}}) do
	pt(mil,"PosteFoco",3,22,3, pos[1],GY+11,pos[2],"Dark grey",Enum.Material.Metal)
	local foco = pt(mil,"Foco",5,3,5, pos[1],GY+23,pos[2],"Bright yellow",Enum.Material.Neon)
	gl(foco,4,55,255,240,180)
end

-- Barriles con fuego
for _,pos in ipairs({{MX-50,MZ+45},{MX+50,MZ+45},{MX,MZ-65}}) do
	pt(mil,"Barril",5,7,5, pos[1],GY+4,pos[2],"Very dark grey",Enum.Material.Metal)
	local bf = pt(mil,"Llama",4,2,4, pos[1],GY+8,pos[2],"Bright orange",Enum.Material.Neon)
	fir(bf,5,5); gl(bf,4,26,255,140,40)
end

pt(mil,"PosteL",3,16,3, MX-12,GY+8,MZ+90,"Dark grey",Enum.Material.Metal)
pt(mil,"PosteR",3,16,3, MX+12,GY+8,MZ+90,"Dark grey",Enum.Material.Metal)
local mls = pt(mil,"Letrero",32,7,2, MX,GY+18,MZ+90,"Dark grey",Enum.Material.Metal)
sgn(mls,"ZONA MILITAR - PROHIBIDO",255,60,60)

print("Campamento Militar generado en ("..MX..","..GY..","..MZ..")")
print("")
print("================================")
print("  PP2 EXTRA - COMPLETO")
print("  1. PuestoCientifico (500,331,-420)")
print("  2. TorreRadio       (780,331,-580)")
print("  3. CampamentoMilitar(900,331,-200)")
print("  Items: #6 lab  #7 radio  #8 bunker")
print("================================")

game.Selection:Set({sci})
ws.CurrentCamera.CFrame = CFrame.new(SX,GY+130,SZ+160)*CFrame.Angles(math.rad(-38),0,0)
