import psycopg2

conn = psycopg2.connect(
    host="localhost",
    database="test_bd_1",  
    user="postgres",          
    password="postgres",      
    port="5432"
)

cur = conn.cursor()

Scada_table = "slabs"

cur.execute("SELECT * FROM slabs ORDER BY id DESC LIMIT 1")



last_row = cur.fetchone()

    
id, Length_slab, Width_slab, Thikness_slab, Temperature_slab, Material_slab,Diametr_roll, Material_roll,is_used = last_row
    
    
print("Последняя запись из таблицы:")
print(f"ID: {id}")
print(f"Колонка 2: {Length_slab}")
print(f"Колонка 3: {Width_slab}")
print(f"Колонка 4: {Thikness_slab}")
print(f"Колонка 5: {Temperature_slab}")
print(f"Колонка 6: {Material_slab}")
print(f"Колонка 7: {Diametr_roll}")
print(f"Колонка 8: {Material_roll}")
print(f"{is_used}")
if is_used == False:
    cur.execute(f"UPDATE public.slabs  SET is_used=true WHERE id = {id};")
    conn.commit()


