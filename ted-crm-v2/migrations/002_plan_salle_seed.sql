-- Plan de salle — 2/2 : reprise du plan actuel (PLAN_SALLE_DEFAUT)
-- Les positions relatives sont conservées ; les dimensions passent en cm réels
-- (une table ronde devient enfin ronde, ce que le rendu en % ne permettait pas).
-- Ne fait rien si des salles existent déjà.
-- À exécuter APRÈS 001, dans Supabase → SQL Editor.

do $$
declare
  id_salle    uuid;
  id_terrasse uuid;
  id_etage    uuid;
begin
  if exists (select 1 from public.salles) then
    raise notice 'Des salles existent déjà — seed ignoré.';
    return;
  end if;

  insert into public.salles (nom, type, ordre, largeur, hauteur)
       values ('Salle', 'salle', 0, 1200, 800) returning id into id_salle;
  insert into public.salles (nom, type, ordre, largeur, hauteur)
       values ('Terrasse', 'terrasse', 1, 1200, 800) returning id into id_terrasse;
  insert into public.salles (nom, type, ordre, largeur, hauteur)
       values ('Étage', 'etage', 2, 1200, 800) returning id into id_etage;

  -- ── Décors ───────────────────────────────────────────────────────────────
  insert into public.decors_salle (salle_id, libelle, type, x, y, largeur, hauteur) values
    (id_salle,    'Bar',         'bar',      36, 64,  144, 400),
    (id_salle,    'Entrée',      'entree',  516, 728, 168,  56),
    (id_salle,    'Escalier',    'escalier',1020, 64, 144, 160),
    (id_salle,    'Cuisine',     'cuisine', 1020, 560, 144, 160),
    (id_terrasse, 'Accès salle', 'entree',   24, 336, 108, 128),
    (id_etage,    'Escalier',    'escalier',  36, 64, 144, 160),
    (id_etage,    'Bar',         'bar',     1008, 64, 156, 256);

  -- ── Tables ───────────────────────────────────────────────────────────────
  -- ronde 2 pers. = 70 cm · carrée 4 pers. = 90×90 · ovale 6 = 180×90
  -- ovale 8 = 220×90 · ovale 10 = 260×90
  insert into public.tables_salle
    (salle_id, libelle, forme, x, y, largeur, hauteur, nb_couverts) values
    (id_salle,    '1',  'ronde',  276,  72,  70,  70, 2),
    (id_salle,    '2',  'ronde',  504,  72,  70,  70, 2),
    (id_salle,    '3',  'ronde',  732,  72,  70,  70, 2),
    (id_salle,    '4',  'carree', 276, 272,  90,  90, 4),
    (id_salle,    '5',  'carree', 540, 272,  90,  90, 4),
    (id_salle,    '6',  'carree', 804, 272,  90,  90, 4),
    (id_salle,    '7',  'ovale',  276, 512, 180,  90, 6),
    (id_salle,    '8',  'ovale',  624, 512, 220,  90, 8),
    (id_terrasse, 'E1', 'ronde',  264, 112,  70,  70, 2),
    (id_terrasse, 'E2', 'ronde',  492, 112,  70,  70, 2),
    (id_terrasse, 'E3', 'ronde',  720, 112,  70,  70, 2),
    (id_terrasse, 'E4', 'carree', 264, 384,  90,  90, 4),
    (id_terrasse, 'E5', 'carree', 528, 384,  90,  90, 4),
    (id_terrasse, 'E6', 'ovale',  768, 392, 180,  90, 6),
    (id_etage,    'A1', 'ovale',  288, 112, 220,  90, 8),
    (id_etage,    'A2', 'ovale',  288, 416, 260,  90, 10),
    (id_etage,    'A3', 'carree', 744, 112,  90,  90, 4),
    (id_etage,    'A4', 'carree', 744, 416,  90,  90, 4);
end $$;

-- ── Reprise des placements existants (reservations.table_plan → table_id) ──
-- Anciens identifiants : T1..T8 dans la salle (libellés « 1 »..« 8 »),
-- E1..E6 et A1..A4 identiques au libellé.
insert into public.reservations_tables (reservation_id, table_id)
select r.id, t.id
  from public.reservations r
  join public.salles s
    on s.nom = case
         when r.table_plan like 'T%' then 'Salle'
         when r.table_plan like 'E%' then 'Terrasse'
         when r.table_plan like 'A%' then 'Étage'
       end
  join public.tables_salle t
    on t.salle_id = s.id
   and t.libelle = case
         when r.table_plan like 'T%' then substring(r.table_plan from 2)
         else r.table_plan
       end
 where r.table_plan is not null
   and r.table_id is null
on conflict do nothing;
