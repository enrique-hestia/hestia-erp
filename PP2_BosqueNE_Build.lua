-- PP2 BOSQUE NE — Build Script v5 (FINAL)
-- Suelo real: Y=331 | Zona: BosqueNE_Zone en Workspace

local ws = game:GetService("Workspace")
local zona = ws:FindFirstChild("BosqueNE_Zone")
assert(zona, "❌ Crea la carpeta BosqueNE_Zone en Workspace primero")

-- Limpiar builds anteriores
for _, c in zona:GetChildren() do
	if c.Name == "PP2_CampamentoForestal" or c.Name == "PP2_FabricaAbandonada" then
		c:Destroy()
	end
end

-- ─── HELPERS ───────────────────────────────────────────────────────────────

local function part(parent, name, sx, sy, sz, x, y, z, color, mat, transp)
	local p = Instance.new("Part")
	p.Name = name
	p.Size = Vector3.new(sx, sy, sz)
	p.CFrame = CFrame.new(x, y, z)
	p.BrickColor = BrickColor.new(color)
	p.Material = mat or Enum.Material.SmoothPlastic
	p.Transparency = transp or 0
	p.Anchored = true
	p.Parent = parent
	return p
end

local function wedge(parent, name, sx, sy, sz, cf, color, mat)
	local p = Instance.new("WedgePart")
	p.Name = name
	p.Size = Vector3.new(sx, sy, sz)
	p.CFrame = cf
	p.BrickColor = BrickColor.new(color)
	p.Material = mat or Enum.Material.SmoothPlastic
	p.Anchored = true
	p.Parent = parent
	return p
end

local function model(parent, name)
	local m = Instance.new("Model")
	m.Name = name
	m.Parent = parent
	return m
end

local function light(p, br, rng, r, g, b)
	local l = Instance.new("PointLight")
	l.Brightness = br; l.Range = rng
	l.Color = Color3.fromRGB(r, g, b)
	l.Parent = p
end

local function prompt(p, txt)
	local pp = Instance.new("ProximityPrompt")
	pp.ActionText = txt
	pp.MaxActivationDistance = 8
	pp.Parent = p
end

local function fire(p, sz, ht)
	local f = Instance.new("Fire")
	f.Size = sz; f.Heat = ht
	f.Parent = p
end

local function smoke(p, sz, op)
	local s = Instance.new("Smoke")
	s.Size = sz; s.Opacity = op
	s.RiseVelocity = 5
	s.Color = Color3.fromRGB(70, 70, 70)
	s.Parent = p
end

local function sign(p, txt, color)
	local sg = Instance.new("SurfaceGui", p)
	sg.Face = Enum.NormalId.Front
	local lb = Instance.new("TextLabel", sg)
	lb.Size = UDim2.new(1,0,1,0)
	lb.BackgroundTransparency = 1
	lb.Text = txt
	lb.TextColor3 = color or Color3.fromRGB(255, 220, 120)
	lb.TextScaled = true
	lb.Font = Enum.Font.GothamBold
end

-- ─── COORDENADAS ───────────────────────────────────────────────────────────

local GY = 331  -- Y del suelo real

-- Campamento centro
local CX, CZ = 700, -200
-- Fábrica centro
local FX, FZ = 920, -450

-- ════════════════════════════════════════════════════════════════════════════
--  CAMPAMENTO FORESTAL
-- ════════════════════════════════════════════════════════════════════════════

local camp = model(zona, "PP2_CampamentoForestal")

-- Claro de tierra
part(camp, "Claro", 200, 2, 200, CX, GY-1, CZ, "Moss", Enum.Material.Grass)

-- ── CABAÑA PRINCIPAL (60×30 base, 26 alto) ─────────────────────────────
local cab = model(camp, "Cabana")

-- Paredes
part(cab, "Frente",  60, 26, 3,  CX,    GY+13, CZ+18, "Reddish brown", Enum.Material.Wood)
part(cab, "Atras",   60, 26, 3,  CX,    GY+13, CZ-18, "Reddish brown", Enum.Material.Wood)
part(cab, "Izq",     3,  26, 36, CX-30, GY+13, CZ,    "Reddish brown", Enum.Material.Wood)
part(cab, "Der",     3,  26, 36, CX+30, GY+13, CZ,    "Reddish brown", Enum.Material.Wood)
-- Piso
part(cab, "Piso",    60, 2,  36, CX,    GY+1,  CZ,    "Dark orange",   Enum.Material.WoodPlanks)
-- Techo a dos aguas
wedge(cab, "Techo_L", 60, 14, 20, CFrame.new(CX, GY+33, CZ-8),                                        "Dark red", Enum.Material.SmoothPlastic)
wedge(cab, "Techo_R", 60, 14, 20, CFrame.new(CX, GY+33, CZ+8) * CFrame.Angles(0, math.pi, 0),        "Dark red", Enum.Material.SmoothPlastic)
-- Chimenea exterior
part(cab, "Chimenea",  6, 40, 6, CX-24, GY+20, CZ-16, "Medium stone grey", Enum.Material.Cobblestone)
-- Puerta (marco)
part(cab, "MarcoPuerta", 8, 12, 1, CX, GY+7, CZ+18.5, "Dark orange", Enum.Material.Wood)
-- Ventanas frente
for _, ox in ipairs({-18, 18}) do
	part(cab, "VMarco"..ox,  10, 8, 1,   CX+ox, GY+16, CZ+18.5, "Dark orange",      Enum.Material.Wood)
	part(cab, "Vidrio"..ox,  8,  6, 0.3, CX+ox, GY+16, CZ+18.6, "Medium stone grey", Enum.Material.Glass, 0.4)
end
-- Ventanas lateral
for _, oz in ipairs({-8, 8}) do
	part(cab, "VLatMarco"..oz,  1, 8, 10,  CX+30.5, GY+16, CZ+oz, "Dark orange",      Enum.Material.Wood)
	part(cab, "VLatVidrio"..oz, 0.3, 6, 8, CX+30.6, GY+16, CZ+oz, "Medium stone grey", Enum.Material.Glass, 0.4)
end
-- Interior
part(cab, "Mesa",    16, 4, 8,  CX-12, GY+3,  CZ-8, "Dark orange",  Enum.Material.WoodPlanks)
part(cab, "Cama1",   10, 4, 18, CX+18, GY+3,  CZ-6, "Bright blue",  Enum.Material.SmoothPlastic)
part(cab, "Cama2",   10, 4, 18, CX+18, GY+3,  CZ+6, "Bright red",   Enum.Material.SmoothPlastic)
part(cab, "Estante", 18, 14, 3, CX-27, GY+8,  CZ+6, "Dark orange",  Enum.Material.WoodPlanks)
part(cab, "Silla",   6,  6, 6,  CX-8,  GY+4,  CZ-4, "Reddish brown",Enum.Material.Wood)
-- Item PP2 #1 en mesa
local i1 = part(cab, "PP2_Item1", 3,3,3, CX-12, GY+8, CZ-8, "Bright yellow", Enum.Material.Neon)
light(i1, 4, 14, 255, 230, 60)
prompt(i1, "Recolectar Batería")

-- ── BODEGA / ALMACÉN (50×24 base, 22 alto) ─────────────────────────────
local bod = model(camp, "Bodega")
local BX, BZ = CX+80, CZ-40

part(bod, "Frente", 50, 22, 3,  BX,    GY+11, BZ+13, "Medium stone grey", Enum.Material.Concrete)
part(bod, "Atras",  50, 22, 3,  BX,    GY+11, BZ-13, "Medium stone grey", Enum.Material.Concrete)
part(bod, "Izq",    3,  22, 26, BX-25, GY+11, BZ,    "Medium stone grey", Enum.Material.Concrete)
part(bod, "Der",    3,  22, 26, BX+25, GY+11, BZ,    "Medium stone grey", Enum.Material.Concrete)
part(bod, "Piso",   50, 2,  26, BX,    GY+1,  BZ,    "Dark grey",         Enum.Material.Concrete)
part(bod, "Techo",  52, 3,  28, BX,    GY+23, BZ,    "Dark grey",         Enum.Material.Metal)
-- Portón
part(bod, "Porton", 16, 18, 1,  BX,    GY+10, BZ+13.5, "Dark grey", Enum.Material.Metal)
-- Estantes
for i = 1, 4 do
	part(bod, "Estante"..i, 3, 18, 22, BX-18+(i*10), GY+10, BZ+2, "Dark orange", Enum.Material.WoodPlanks)
end
-- Cajas
part(bod, "Caja1", 6,6,6, BX+20, GY+4,  BZ-8, "Reddish brown", Enum.Material.WoodPlanks)
part(bod, "Caja2", 6,6,6, BX+20, GY+10, BZ-8, "Reddish brown", Enum.Material.WoodPlanks)
part(bod, "Caja3", 6,6,6, BX+14, GY+4,  BZ-8, "Reddish brown", Enum.Material.WoodPlanks)
-- Item PP2 #2
local i2 = part(bod, "PP2_Item2", 3,3,3, BX+20, GY+15, BZ-8, "Bright yellow", Enum.Material.Neon)
light(i2, 3, 12, 255, 220, 50)
prompt(i2, "Recolectar Componente")

-- ── TORRE DE OBSERVACIÓN (70 studs de alto) ───────────────────────────
local tor = model(camp, "TorreObservacion")
local TX, TZ = CX-80, CZ-60

-- 4 columnas
for _, ox in ipairs({-5,5}) do
	for _, oz in ipairs({-5,5}) do
		part(tor, "Col"..ox..oz, 3, 75, 3, TX+ox, GY+37, TZ+oz, "Reddish brown", Enum.Material.Wood)
	end
end
-- 4 plataformas + barandales
for i, h in ipairs({20, 38, 56, 74}) do
	part(tor, "Plat"..i, 14, 2, 14, TX, GY+h, TZ, "Dark orange", Enum.Material.WoodPlanks)
	-- barandales
	part(tor, "RL"..i, 2, 4, 14, TX-7, GY+h+2, TZ, "Reddish brown", Enum.Material.Wood)
	part(tor, "RR"..i, 2, 4, 14, TX+7, GY+h+2, TZ, "Reddish brown", Enum.Material.Wood)
	part(tor, "RN"..i, 14, 4, 2, TX, GY+h+2, TZ-7, "Reddish brown", Enum.Material.Wood)
	part(tor, "RS"..i, 14, 4, 2, TX, GY+h+2, TZ+7, "Reddish brown", Enum.Material.Wood)
	-- escalera entre plataformas
	if i < 4 then
		wedge(tor, "Esc"..i, 4, 18, 10,
			CFrame.new(TX+5, GY+h+10, TZ) * CFrame.Angles(0, math.pi, 0),
			"Dark orange", Enum.Material.WoodPlanks)
	end
end
-- Cima con fanal y PP2 Core
part(tor, "Fanal", 10, 10, 10, TX, GY+82, TZ, "Medium stone grey", Enum.Material.Glass, 0.4)
local pp2 = part(tor, "PP2_Core", 5,5,5, TX, GY+79, TZ, "Bright yellow", Enum.Material.Neon)
light(pp2, 10, 50, 255, 220, 50)
prompt(pp2, "Activar PowerPackage 2")

-- ── CERCA PERIMETRAL ──────────────────────────────────────────────────
local cerca = model(camp, "Cerca")
local cposts = {
	{CX-100,CZ-100},{CX-50,CZ-100},{CX,CZ-100},{CX+50,CZ-100},{CX+100,CZ-100},
	{CX+100,CZ-50},{CX+100,CZ},{CX+100,CZ+50},{CX+100,CZ+100},
	{CX+50,CZ+100},{CX,CZ+100},{CX-50,CZ+100},{CX-100,CZ+100},
	{CX-100,CZ+50},{CX-100,CZ},{CX-100,CZ-50},
}
for i, p in ipairs(cposts) do
	part(cerca, "Post"..i, 3,12,3, p[1], GY+6, p[2], "Reddish brown", Enum.Material.Wood)
end
for i = 1, #cposts do
	local a, b = cposts[i], cposts[i==#cposts and 1 or i+1]
	local mx, mz = (a[1]+b[1])/2, (a[2]+b[2])/2
	local dx, dz = b[1]-a[1], b[2]-a[2]
	local len = math.sqrt(dx*dx+dz*dz)
	local ang = math.atan2(dz, dx)
	for _, oy in ipairs({4, 9}) do
		local tb = Instance.new("Part")
		tb.Size = Vector3.new(len, 2, 2)
		tb.CFrame = CFrame.new(mx, GY+oy, mz) * CFrame.Angles(0, -ang, 0)
		tb.BrickColor = BrickColor.new("Reddish brown")
		tb.Material = Enum.Material.Wood
		tb.Anchored = true
		tb.Parent = cerca
	end
end

-- ── HOGUERA ────────────────────────────────────────────────────────────
local hog = model(camp, "Hoguera")
part(hog, "Piedras", 8, 2, 8, CX+15, GY+1, CZ+50, "Medium stone grey", Enum.Material.Cobblestone)
local fueg = part(hog, "Fuego", 4, 3, 4, CX+15, GY+3, CZ+50, "Bright orange", Enum.Material.Neon)
fire(fueg, 8, 6)
light(fueg, 5, 35, 255, 140, 40)

-- ── CARTEL ENTRADA ──────────────────────────────────────────────────────
part(camp, "PosteL", 3,16,3, CX-12, GY+8,  CZ+100, "Reddish brown", Enum.Material.Wood)
part(camp, "PosteR", 3,16,3, CX+12, GY+8,  CZ+100, "Reddish brown", Enum.Material.Wood)
local cart = part(camp, "Cartel", 28, 6, 2, CX, GY+18, CZ+100, "Reddish brown", Enum.Material.WoodPlanks)
sign(cart, "ESTACIÓN GUARDABOSQUES")

-- ── ITEM ESCONDIDO BOSQUE ──────────────────────────────────────────────
local i3 = part(camp, "PP2_Item3", 3,3,3, CX-80, GY+2, CZ+70, "Bright yellow", Enum.Material.Neon)
light(i3, 3, 12, 255, 220, 50)
prompt(i3, "Recolectar Pieza Oculta")

print("✅ Campamento Forestal generado en ("..CX..","..GY..","..CZ..")")

-- ════════════════════════════════════════════════════════════════════════════
--  FÁBRICA ABANDONADA
-- ════════════════════════════════════════════════════════════════════════════

local fab = model(zona, "PP2_FabricaAbandonada")

-- Suelo
part(fab, "Suelo", 240, 2, 160, FX, GY-1, FZ, "Dark grey", Enum.Material.Concrete)

-- ── EDIFICIO PRINCIPAL (100×36 base, 34 alto) ──────────────────────────
local edi = model(fab, "EdificioPrincipal")

part(edi, "Norte",   100, 34, 4, FX,    GY+17, FZ-30, "Medium stone grey", Enum.Material.Concrete)
part(edi, "Sur",     100, 34, 4, FX,    GY+17, FZ+30, "Medium stone grey", Enum.Material.Concrete)
part(edi, "Este",    4,   34, 60, FX+50, GY+17, FZ,   "Medium stone grey", Enum.Material.Concrete)
part(edi, "Oeste",   4,   34, 60, FX-50, GY+17, FZ,   "Medium stone grey", Enum.Material.Concrete)
part(edi, "Piso",    100, 2,  60, FX,    GY+1,  FZ,   "Dark grey",         Enum.Material.Concrete)
-- Techo con hueco
part(edi, "Techo_L", 40, 4, 60, FX-30, GY+35, FZ, "Dark grey", Enum.Material.Metal)
part(edi, "Techo_R", 40, 4, 60, FX+30, GY+35, FZ, "Dark grey", Enum.Material.Metal)
-- Claraboya (hueco con vidrio)
part(edi, "Claraboya", 18, 2, 24, FX, GY+36, FZ, "Medium stone grey", Enum.Material.Glass, 0.5)

-- Ventanas rotas norte
for _, ox in ipairs({-35,-15,15,35}) do
	part(edi, "VN"..ox, 10, 10, 1, FX+ox, GY+22, FZ-30.5, "Dark grey", Enum.Material.Metal)
	part(edi, "VVN"..ox, 7, 7, 0.4, FX+ox-1, FZ-30.6, GY+20, "Medium stone grey", Enum.Material.Glass, 0.5)
end
-- Ventanas sur
for _, ox in ipairs({-35,-15,15,35}) do
	part(edi, "VS"..ox, 10, 10, 1, FX+ox, GY+22, FZ+30.5, "Dark grey", Enum.Material.Metal)
end

-- Portón principal (semiabierto)
part(edi, "PortonL", 14, 24, 2, FX-14, GY+12, FZ+30.5, "Dark grey", Enum.Material.Metal)
local portR = Instance.new("Part")
portR.Size = Vector3.new(10, 24, 2)
portR.CFrame = CFrame.new(FX+10, GY+12, FZ+30) * CFrame.Angles(0, 0, math.rad(50))
portR.BrickColor = BrickColor.new("Dark grey")
portR.Material = Enum.Material.Metal
portR.Anchored = true
portR.Parent = edi

-- Maquinaria interior
-- Tanques industriales
for i, ox in ipairs({-35,-18,0,18}) do
	local tank = part(edi, "Tanque"..i, 10, 22, 10, FX+ox, GY+12, FZ-12, "Dark grey", Enum.Material.Metal)
	if i <= 2 then
		local top = part(edi, "TanqueTop"..i, 10,3,10, FX+ox, GY+24, FZ-12, "Very dark grey", Enum.Material.Metal)
		smoke(top, 1.5, 0.8)
	end
end
-- Tuberías horizontales
part(edi, "Tub1", 55, 3, 3, FX-8, GY+24, FZ-12, "Dark grey", Enum.Material.Metal)
part(edi, "Tub2", 3, 16, 3, FX+27, GY+16, FZ-12, "Dark grey", Enum.Material.Metal)

-- Cinta transportadora
part(edi, "Cinta",   70, 5, 10, FX+5, GY+6, FZ+12, "Very dark grey", Enum.Material.Metal)
part(edi, "Rod1",    3,  8, 12, FX-30, GY+6, FZ+12, "Dark grey", Enum.Material.Metal)
part(edi, "Rod2",    3,  8, 12, FX+40, GY+6, FZ+12, "Dark grey", Enum.Material.Metal)
wedge(edi, "CintaRota", 22, 4, 10,
	CFrame.new(FX+34, GY+2, FZ+12) * CFrame.Angles(0, 0, math.rad(-25)),
	"Very dark grey", Enum.Material.Metal)

-- Cajas apiladas
for i = 1, 4 do
	for j = 1, 2 do
		part(edi, "Caja"..i.."_"..j, 8,8,8,
			FX+38, GY+5+(j-1)*8, FZ-16+(i-1)*10, "Reddish brown", Enum.Material.WoodPlanks)
	end
end

-- ── CHIMENEAS EXTERIORES ───────────────────────────────────────────────
for _, ox in ipairs({-24, 24}) do
	part(fab, "Chimenea"..ox, 10, 70, 10, FX+ox, GY+35, FZ-30, "Dark grey", Enum.Material.Concrete)
	local ctop = part(fab, "ChimTop"..ox, 10,3,10, FX+ox, GY+71, FZ-30, "Very dark grey", Enum.Material.Metal)
	smoke(ctop, 2, 0.9)
	part(fab, "BandaRoja"..ox, 12,4,12, FX+ox, GY+58, FZ-30, "Bright red", Enum.Material.SmoothPlastic)
end

-- ── EDIFICIO OFICINAS (50×20 base, 20 alto) ────────────────────────────
local ofi = model(fab, "Oficinas")
local OX, OZ = FX+100, FZ-22

part(ofi, "Norte", 50, 20, 3, OX,    GY+10, OZ-18, "Sand green", Enum.Material.Concrete)
part(ofi, "Sur",   50, 20, 3, OX,    GY+10, OZ+18, "Sand green", Enum.Material.Concrete)
part(ofi, "Este",  3,  20, 36, OX+25, GY+10, OZ,   "Sand green", Enum.Material.Concrete)
part(ofi, "Oeste", 3,  20, 36, OX-25, GY+10, OZ,   "Sand green", Enum.Material.Concrete)
part(ofi, "Piso",  50, 2,  36, OX,    GY+1,  OZ,   "Dark grey",  Enum.Material.Concrete)
part(ofi, "Techo", 52, 3,  38, OX,    GY+21, OZ,   "Sand green", Enum.Material.Metal)
-- Ventanas
for _, ox in ipairs({-16, 0, 16}) do
	part(ofi, "Vent"..ox, 10, 7, 0.5, OX+ox, GY+13, OZ+18.2, "Medium stone grey", Enum.Material.Glass, 0.35)
end
-- Interior
for i = 1, 5 do
	part(ofi, "Escritorio"..i, 10,4,6, OX-20+(i*9), GY+3, OZ-6, "Reddish brown", Enum.Material.WoodPlanks)
end
part(ofi, "ArchiveroP", 5,14,5, OX+20, GY+4, OZ+12, "Dark grey", Enum.Material.Metal)
local archC = Instance.new("Part")
archC.Size = Vector3.new(5,14,5)
archC.CFrame = CFrame.new(OX+14, GY+4, OZ+12) * CFrame.Angles(0,0,math.rad(90))
archC.BrickColor = BrickColor.new("Dark grey")
archC.Material = Enum.Material.Metal
archC.Anchored = true
archC.Parent = ofi
-- Item PP2 #4
local i4 = part(ofi, "PP2_Item4", 3,3,3, OX, GY+7, OZ-6, "Bright yellow", Enum.Material.Neon)
light(i4, 3, 12, 255, 220, 50)
prompt(i4, "Recolectar Documento")

-- ── PASARELA entre edificios ──────────────────────────────────────────
local pas = model(fab, "Pasarela")
part(pas, "SopL",  4,22,4, FX+50, GY+11, FZ-10, "Dark grey", Enum.Material.Metal)
part(pas, "SopR",  4,22,4, FX+75, GY+11, FZ-10, "Dark grey", Enum.Material.Metal)
part(pas, "Plat",  60, 3, 8, FX+75, GY+23, FZ-18, "Dark grey", Enum.Material.Metal)
part(pas, "RailN", 60, 4, 2, FX+75, GY+26, FZ-22, "Dark grey", Enum.Material.Metal)
part(pas, "RailS", 60, 4, 2, FX+75, GY+26, FZ-14, "Dark grey", Enum.Material.Metal)
-- Item PP2 #5
local i5 = part(pas, "PP2_Item5", 3,3,3, FX+75, GY+26, FZ-18, "Bright yellow", Enum.Material.Neon)
light(i5, 3, 12, 255, 220, 50)
prompt(i5, "Recolectar Engrane")

-- ── BARRILES DE FUEGO ─────────────────────────────────────────────────
for _, pos in ipairs({{FX-55,FZ+15},{FX-55,FZ-5},{FX+55,FZ+20},{FX-55,FZ-25}}) do
	part(fab, "Barril", 5,7,5, pos[1], GY+4, pos[2], "Dark grey", Enum.Material.Metal)
	local ft = part(fab, "Llama", 4,2,4, pos[1], GY+8, pos[2], "Bright orange", Enum.Material.Neon)
	fire(ft, 5, 5)
	light(ft, 4, 22, 255, 140, 40)
end

-- ── VALLA METÁLICA ────────────────────────────────────────────────────
local val = model(fab, "Valla")
local vp = {
	{FX-120,FZ-80},{FX-60,FZ-80},{FX,FZ-80},{FX+60,FZ-80},{FX+120,FZ-80},
	{FX+120,FZ},{FX+120,FZ+80},
	{FX+60,FZ+80},{FX,FZ+80},{FX-60,FZ+80},{FX-120,FZ+80},
	{FX-120,FZ},
}
for i, p in ipairs(vp) do
	part(val, "VP"..i, 3,14,3, p[1], GY+7, p[2], "Dark grey", Enum.Material.Metal)
end
for i = 1, #vp do
	local a, b = vp[i], vp[i==#vp and 1 or i+1]
	local mx, mz = (a[1]+b[1])/2, (a[2]+b[2])/2
	local dx, dz = b[1]-a[1], b[2]-a[2]
	local len = math.sqrt(dx*dx+dz*dz)
	local ang = math.atan2(dz, dx)
	for _, oy in ipairs({4, 9}) do
		local vr = Instance.new("Part")
		vr.Size = Vector3.new(len, 1.5, 1.5)
		vr.CFrame = CFrame.new(mx, GY+oy, mz) * CFrame.Angles(0, -ang, 0)
		vr.BrickColor = BrickColor.new("Dark grey")
		vr.Material = Enum.Material.Metal
		vr.Anchored = true
		vr.Parent = val
	end
end

-- ── LETRERO FÁBRICA ───────────────────────────────────────────────────
local let = part(fab, "Letrero", 36,8,2, FX, GY+38, FZ+31, "Dark grey", Enum.Material.Metal)
sign(let, "⚠ ZONA RESTRINGIDA", Color3.fromRGB(255,60,60))

print("✅ Fábrica Abandonada generada en ("..FX..","..GY..","..FZ..")")
print("")
print("════════════════════════════════")
print("  PP2 BOSQUE NE — COMPLETO")
print("  Campamento: ("..CX..","..GY..","..CZ..")")
print("  Fábrica:    ("..FX..","..GY..","..FZ..")")
print("  Items PP2:  5 cubos amarillos Neon")
print("════════════════════════════════")

-- Seleccionar y volar cámara al campamento
game.Selection:Set({camp})
ws.CurrentCamera.CFrame = CFrame.new(CX, GY+120, CZ+150) * CFrame.Angles(math.rad(-35), 0, 0)
