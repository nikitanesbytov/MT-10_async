import threading
import time

class Server:
    def fast_monitor(self):
        while True:
            print("Быстрый мониторинг (1 сек)")
            # Ваша логика здесь
            time.sleep(1)
    
    def slow_monitor(self):
        while True:
            print("Медленный мониторинг (3 сек)")
            # Ваша логика здесь
            time.sleep(3)

# Создание и запуск потоков
server = Server()

fast_thread = threading.Thread(target=server.fast_monitor, daemon=True)
slow_thread = threading.Thread(target=server.slow_monitor, daemon=True)

fast_thread.start()
slow_thread.start()

# Главный поток продолжает работу
print("Сервер запущен с двумя мониторами...")
time.sleep(9)
print("Завершение работы...")