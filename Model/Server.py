import asyncio
from pymodbus.server import StartAsyncTcpServer
from pymodbus.datastore import ModbusSequentialDataBlock
from pymodbus.datastore import ModbusSlaveContext, ModbusServerContext
import struct
import asyncpg
from RollingMillSimulator import RollingMillSimulator

def float_to_regs(value):
    """Преобразует float в два WORD регистра (big-endian)"""
    b = struct.pack('>f', float(value))
    return [int.from_bytes(b[:2], 'big'), int.from_bytes(b[2:], 'big')]

def regs_to_float(reg1, reg2):
    """Преобразует два WORD регистра обратно в float (big-endian)"""
    b1 = (reg1 >> 8) & 0xFF
    b2 = reg1 & 0xFF
    b3 = (reg2 >> 8) & 0xFF
    b4 = reg2 & 0xFF
    return struct.unpack('>f', bytes([b1, b2, b3, b4]))[0]

class AsyncModbusServer:
    def __init__(self):  
        total_registers = 31
        initial_values = [0] * total_registers
        self.hr_data_combined = ModbusSequentialDataBlock(1, initial_values)
        store = ModbusSlaveContext(hr=self.hr_data_combined)
        self.context = ModbusServerContext(slaves=store, single=True)
        self.stop_monitoring = False
        self.simulator = None
        self.initialized = False
        self.counter = 0
        self.counter2 = 0
        self.nex_idx = 0
        self.prev_total_steps = 0
        self.simulation_lock = asyncio.Lock()
        self.simulation_in_progress = False

    async def alarm_stop(self, diff):
        """Асинхронное выполнение последовательности аварийной остановки"""
        async with self.simulation_lock:
            self.simulation_in_progress = True
            try:
                alarm_data = self.simulator.Alarm_stop()
                await self._write_alarm_data_to_registers(alarm_data, diff)
            finally:
                self.simulation_in_progress = False

    async def _write_alarm_data_to_registers(self, alarm_data, diff):
        """Асинхронно записывает данные аварийной остановки в регистры"""
        total_steps = len(alarm_data['Time']) 
        self.nex_idx += diff
        
        while self.nex_idx != total_steps and not self.stop_monitoring:      
            await self._write_single_step_to_registers_sync(alarm_data, self.nex_idx)
            self.nex_idx += 1
            await asyncio.sleep(0.1)
            
    async def update_simulation_registers(self, sim_data, idx):
        """Асинхронное обновление регистров симуляции"""   
        keys = [
            'Pyro1', 'Pyro2', 'Pressure', 'Gap', 'VRPM', 'V0RPM', 'V1RPM',
            'Moment', 'Power'
        ]
        regs = []
        for k in keys:
            v = sim_data[k][idx] if isinstance(sim_data[k], list) else sim_data[k]
            regs.extend(float_to_regs(v))
        
        flags = 0
        StartCap_val = sim_data['StartCap'][idx] if isinstance(sim_data['StartCap'], list) else sim_data['StartCap']
        EndCap_val = sim_data['EndCap'][idx] if isinstance(sim_data['EndCap'], list) else sim_data['EndCap']
        Gap_feedback_val = sim_data['Gap_feedback'][idx] if isinstance(sim_data['Gap_feedback'], list) else sim_data['Gap_feedback']
        Speed_feedback_val = sim_data['Speed_feedback'][idx] if isinstance(sim_data['Speed_feedback'], list) else sim_data['Speed_feedback']
        
        if StartCap_val:
            flags |= 0x01
        if EndCap_val:
            flags |= 0x02
        if Gap_feedback_val:
            flags |= 0x04
        if Speed_feedback_val:
            flags |= 0x08
        
        self.hr_data_combined.setValues(12, regs)  
        self.hr_data_combined.setValues(30, [flags])

    async def start_init_from_registers(self):
        """Асинхронная инициализация из БД"""
        async with self.simulation_lock:
            self.simulation_in_progress = True
            try:
                reg30 = self.hr_data_combined.getValues(30, 1)[0]
                reg30 = 0
                new_reg30 = 0
                
                self.hr_data_combined.setValues(12, [0] * 20)
                gap_regs = float_to_regs(350)
                self.hr_data_combined.setValues(18, gap_regs)
                while not self.stop_monitoring:
                    try:
                        conn = await asyncio.wait_for(
                            asyncpg.connect(
                                host="localhost",
                                database="postgres",
                                user="postgres",
                                password="postgres",
                                port="5432"
                            ), timeout=1.0
                        )
                        
                        # Проверяем количество записей и удаляем старые
                        count_result = await conn.fetchrow("SELECT COUNT(*) as count FROM slabs")
                        if count_result and count_result['count'] > 10:
                            # Оставляем только 3 последние записи
                            await conn.execute("""
                                DELETE FROM slabs 
                                WHERE id NOT IN (
                                    SELECT id FROM slabs 
                                    ORDER BY id DESC 
                                    LIMIT 3
                                )
                            """)
                        
                        last_row = await conn.fetchrow("SELECT * FROM slabs ORDER BY id DESC LIMIT 1")
                        if last_row and not last_row['is_used']:
                            sim = RollingMillSimulator(
                                L=0, b=0, h_0=0, S=0, StartTemp=0,
                                DV=0, MV=0, MS=0, OutTemp=0, DR=0, SteelGrade=0,
                                V0=0, V1=0, VS=0, Dir_of_rot=0,
                                d1=0, d2=0, d=0, V_Valk_Per=0, StartS=350
                            )
                            
                            ms_clean = (last_row['material_slab'] or "").replace(' ', '')
                            sim.Init(
                                Length_slab=last_row['length_slab'],
                                Width_slab=last_row['width_slab'],
                                Thikness_slab=last_row['thikness_slab'],
                                Temperature_slab=last_row['temperature_slab'],
                                Material_slab=ms_clean,
                                Diametr_roll=last_row['diametr_roll'],
                                Material_roll=last_row['material_roll']
                            )
                            
                            self.simulator = sim
                            self.initialized = True
                            new_reg30 = reg30 | 0x10
                            self.hr_data_combined.setValues(30, [new_reg30])

                            await conn.execute(f"UPDATE public.slabs SET is_used=true WHERE id = {last_row['id']}")
                            await conn.close()
                            
                            self.nex_idx = 0
                            break
                    except asyncio.TimeoutError:
                        continue
            finally:
                self.simulation_in_progress = False

    async def run_server(self, IP, port):
        """Асинхронный запуск Modbus сервера"""
        try:
            await StartAsyncTcpServer(context=self.context, address=(IP, port))
        finally:
            self.stop_monitoring = True

    async def write_simulation_data_to_registers(self, sim_data):
        """Запись данных симуляции в регистры"""
        async with self.simulation_lock:
            self.simulation_in_progress = True
            try:
                total_steps = len(sim_data['Time'])    
                diff = total_steps - self.prev_total_steps
                self.prev_total_steps = total_steps
                while self.nex_idx != total_steps and not self.stop_monitoring:      
                    self._write_single_step_to_registers_sync(sim_data, self.nex_idx)
                    self.nex_idx += 1
                    diff -= 1
        
                    regs = self.hr_data_combined.getValues(1, 11)
                    reg8 = regs[8]
                    Alarm = bool(reg8 & 0x08)
                    Alarm_stop = bool(reg8 & 0x01)
                    
                    if Alarm:
                        await self.alarm_stop(diff)
                        self.initialized = False
                        return
                    if Alarm_stop:
                        await self.start_init_from_registers()
                        self.initialized = False
                        return
                    await asyncio.sleep(0.1)  
            finally:
                self.simulation_in_progress = False

    def _write_single_step_to_registers_sync(self, sim_data, idx):
        """Синхронно записывает данные одного шага симуляции в регистры"""
        keys = [
            'Pyro1', 'Pyro2', 'Pressure', 'Gap', 'VRPM', 'V0RPM', 'V1RPM',
            'Moment', 'Power'
        ]
        regs = []
        for k in keys:
            v = sim_data[k][idx] if isinstance(sim_data[k], list) else sim_data[k]
            regs.extend(float_to_regs(v))
        
        self.hr_data_combined.setValues(12, regs)  
        
        flags = 0
        StartCap_val = sim_data['StartCap'][idx] if isinstance(sim_data['StartCap'], list) else sim_data['StartCap']
        EndCap_val = sim_data['EndCap'][idx] if isinstance(sim_data['EndCap'], list) else sim_data['EndCap']
        Gap_feedback_val = sim_data['Gap_feedback'][idx] if isinstance(sim_data['Gap_feedback'], list) else sim_data['Gap_feedback']
        Speed_feedback_val = sim_data['Speed_feedback'][idx] if isinstance(sim_data['Speed_feedback'], list) else sim_data['Speed_feedback']
        
        if StartCap_val:
            flags |= 0x01
        if EndCap_val:
            flags |= 0x02
        if Gap_feedback_val:
            flags |= 0x04
        if Speed_feedback_val:
            flags |= 0x08
        
        self.hr_data_combined.setValues(30, [flags])
         
    async def monitor_registers(self):
        """Мониторинг регистров с проверкой блокировки"""
        while not self.stop_monitoring:
            if self.simulation_in_progress:
                await asyncio.sleep(0.1)
                continue
                
            if self.simulation_lock.locked():
                await asyncio.sleep(0.1)
                continue
            if not self.initialized:
                await asyncio.sleep(0.1)
                continue
            
            regs = self.hr_data_combined.getValues(1, 9)
            reg8 = regs[8]
    
            Alarm_stop = bool(reg8 & 0x01)
            Start = bool(reg8 & 0x10)
            Start_Gap = bool(reg8 & 0x20)
            Start_Accel = bool(reg8 & 0x40)
            Start_Roll = bool(reg8 & 0x80)
            
            if Alarm_stop:
                self.initialized = False
                await self.start_init_from_registers()
                continue
                
            if Start and not self.simulation_in_progress:
                if Start_Gap and self.counter == 0 and self.counter2 < 2:
                    Roll_pos = regs_to_float(regs[2], regs[3])
                    Dir_of_rot_valk = bool(reg8 & 0x02)
                    sim_result = self.simulator._Gap_Valk_(Roll_pos, Dir_of_rot_valk)
                    await self.write_simulation_data_to_registers(sim_result)
                    self.counter = 1
                    self.counter2 += 1

                if Start_Accel and self.counter == 1 and self.counter2 < 2:
                    Num_of_revol_rolls = regs_to_float(regs[0], regs[1])
                    Dir_of_rot_rolg = bool(reg8 & 0x02)
                    sim_result = self.simulator._Accel_Valk_(Num_of_revol_rolls, Dir_of_rot_rolg, Dir_of_rot_rolg)
                    await self.write_simulation_data_to_registers(sim_result)
                    self.counter = 2
                    self.counter2 += 1

                if Start_Roll and self.counter == 2 and self.counter2 <= 2:
                    Num_of_revol_0rollg = regs_to_float(regs[4], regs[5])
                    Num_of_revol_1rollg = regs_to_float(regs[6], regs[7])
                    Dir_of_rot = bool(reg8 & 0x02)
                    sim_result = self.simulator._Approching_to_Roll_(
                        Dir_of_rot,
                        Num_of_revol_0rollg,
                        Num_of_revol_1rollg,
                    )
                    await self.write_simulation_data_to_registers(sim_result)
                    sim_result = self.simulator._simulate_rolling_pass()
                    await self.write_simulation_data_to_registers(sim_result)
                    sim_result = self.simulator._simulate_exit_from_rolls()
                    await self.write_simulation_data_to_registers(sim_result)
                    self.counter2 += 1
            else:
                self.counter = 0
                self.counter2 = 0
            await asyncio.sleep(0.1)
          

async def main():
    server = AsyncModbusServer()
    
    server_task = asyncio.create_task(server.run_server("192.168.0.99", 55000))
    
    await asyncio.sleep(0.1)
    
    await server.start_init_from_registers()
    
    await asyncio.gather(
        server.monitor_registers(),
        server_task 
    )

if __name__ == "__main__":
    asyncio.run(main())