import json
import os
from datetime import datetime


class CinemaSystem:
    def __init__(self):
        self.theatres_dir = "theatres"
        if not os.path.exists(self.theatres_dir):
            os.makedirs(self.theatres_dir)

    def add_theatre(self, name):
        theatre_data = {
            "name": name,
            "halls": []
        }
        filename = os.path.join(self.theatres_dir, f"{name}.json")

        if os.path.exists(filename):
            print(f"Кинотеатр '{name}' уже существует!")
            return False

        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(theatre_data, f, ensure_ascii=False, indent=2)
        print(f"Кинотеатр '{name}' успешно добавлен!")
        return True

    def get_theatre(self, name):
        filename = os.path.join(self.theatres_dir, f"{name}.json")
        if not os.path.exists(filename):
            return None

        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)

    def save_theatre(self, name, data):
        filename = os.path.join(self.theatres_dir, f"{name}.json")
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def list_theatres(self):
        files = os.listdir(self.theatres_dir)
        theatres = [f.replace('.json', '') for f in files if f.endswith('.json')]
        return theatres

    def add_hall(self, theatre_name, hall_number, rows, seats_per_row):
        theatre = self.get_theatre(theatre_name)
        if not theatre:
            print(f"Кинотеатр '{theatre_name}' не найден!")
            return False

        for hall in theatre["halls"]:
            if hall["number"] == hall_number:
                print(f"Зал №{hall_number} уже существует в кинотеатре '{theatre_name}'!")
                return False

        hall_data = {
            "number": hall_number,
            "rows": rows,
            "seats_per_row": seats_per_row,
            "sessions": []
        }

        theatre["halls"].append(hall_data)
        self.save_theatre(theatre_name, theatre)
        print(f"Зал №{hall_number} добавлен в кинотеатр '{theatre_name}'!")
        return True

    def create_session(self, theatre_name, hall_number, movie_name, start_time, duration):
        theatre = self.get_theatre(theatre_name)
        if not theatre:
            print(f"Кинотеатр '{theatre_name}' не найден!")
            return False

        hall = None
        for h in theatre["halls"]:
            if h["number"] == hall_number:
                hall = h
                break

        if not hall:
            print(f"Зал №{hall_number} не найден в кинотеатре '{theatre_name}'!")
            return False

        seats = [[False for _ in range(hall["seats_per_row"])] for _ in range(hall["rows"])]

        session_data = {
            "movie": movie_name,
            "start_time": start_time,
            "duration": duration,
            "seats": seats
        }

        hall["sessions"].append(session_data)
        self.save_theatre(theatre_name, theatre)
        print(f"Сеанс фильма '{movie_name}' создан на {start_time}!")
        return True

    def sell_ticket(self, theatre_name, hall_number, session_index, row, seat):
        theatre = self.get_theatre(theatre_name)
        if not theatre:
            print(f"Кинотеатр '{theatre_name}' не найден!")
            return False

        hall = None
        for h in theatre["halls"]:
            if h["number"] == hall_number:
                hall = h
                break

        if not hall:
            print(f"Зал №{hall_number} не найден!")
            return False

        if session_index >= len(hall["sessions"]):
            print("Сеанс не найден!")
            return False

        session = hall["sessions"][session_index]

        if row < 0 or row >= hall["rows"]:
            print(f"Ряд {row + 1} не существует!")
            return False

        if seat < 0 or seat >= hall["seats_per_row"]:
            print(f"Место {seat + 1} не существует!")
            return False

        if session["seats"][row][seat]:
            print(f"Место {row + 1}-{seat + 1} уже занято!")
            return False

        session["seats"][row][seat] = True
        self.save_theatre(theatre_name, theatre)
        print(f"Билет продан! Кинотеатр: {theatre_name}, Зал: {hall_number}, "
              f"Фильм: {session['movie']}, Время: {session['start_time']}, "
              f"Место: Ряд {row + 1}, Место {seat + 1}")
        return True

    def find_nearest_session(self, movie_name):
        current_time = datetime.now()
        nearest_session = None
        nearest_time = None
        nearest_info = None

        for theatre_name in self.list_theatres():
            theatre = self.get_theatre(theatre_name)

            for hall in theatre["halls"]:
                for session_index, session in enumerate(hall["sessions"]):
                    if session["movie"] == movie_name:
                        has_free_seats = False
                        for row in session["seats"]:
                            if False in row:
                                has_free_seats = True
                                break

                        if has_free_seats:
                            try:
                                session_time = datetime.strptime(session["start_time"], "%Y-%m-%d %H:%M")

                                if session_time > current_time:
                                    if nearest_time is None or session_time < nearest_time:
                                        nearest_time = session_time
                                        nearest_session = session
                                        nearest_info = {
                                            "theatre": theatre_name,
                                            "hall": hall["number"],
                                            "session_index": session_index
                                        }
                            except ValueError:
                                pass

        if nearest_session:
            print(f"\nБлижайший сеанс фильма '{movie_name}':")
            print(f"Кинотеатр: {nearest_info['theatre']}")
            print(f"Зал: {nearest_info['hall']}")
            print(f"Время: {nearest_session['start_time']}")
            print(f"Длительность: {nearest_session['duration']} мин")
            return nearest_info
        else:
            print(f"Сеансы фильма '{movie_name}' со свободными местами не найдены.")
            return None

    def print_hall_plan(self, theatre_name, hall_number, session_index):
        theatre = self.get_theatre(theatre_name)
        if not theatre:
            print(f"Кинотеатр '{theatre_name}' не найден!")
            return False

        hall = None
        for h in theatre["halls"]:
            if h["number"] == hall_number:
                hall = h
                break

        if not hall:
            print(f"Зал №{hall_number} не найден!")
            return False

        if session_index >= len(hall["sessions"]):
            print("Сеанс не найден!")
            return False

        session = hall["sessions"][session_index]

        print(f"\n{'=' * 60}")
        print(f"ПЛАН ЗАЛА - Кинотеатр: {theatre_name}, Зал: {hall_number}")
        print(f"Фильм: {session['movie']}, Время: {session['start_time']}")
        print(f"{'=' * 60}")
        print("\n                    ЭКРАН")
        print(f"{'-' * 60}\n")

        print("     ", end="")
        for seat_num in range(hall["seats_per_row"]):
            print(f"{seat_num + 1:3}", end=" ")
        print("\n")

        free_count = 0
        occupied_count = 0

        for row_num, row in enumerate(session["seats"]):
            print(f"Ряд {row_num + 1:2} ", end="")
            for seat in row:
                if seat:
                    print(" X ", end=" ")
                    occupied_count += 1
                else:
                    print(" O ", end=" ")
                    free_count += 1
            print()

        print(f"\n{'=' * 60}")
        print(f"Обозначения: O - свободно, X - занято")
        print(f"Свободных мест: {free_count}, Занятых мест: {occupied_count}")
        print(f"{'=' * 60}\n")
        return True


def main():
    system = CinemaSystem()

    print("=" * 60)
    print("БИЛЕТНАЯ СИСТЕМА КИНОТЕАТРОВ".center(60))
    print("=" * 60)

    while True:
        print("\n" + "=" * 60)
        print("ГЛАВНОЕ МЕНЮ")
        print("=" * 60)
        print("1. Добавить кинотеатр")
        print("2. Добавить зал в кинотеатр")
        print("3. Создать сеанс")
        print("4. Продать билет")
        print("5. Найти ближайший сеанс фильма")
        print("6. Показать план зала")
        print("7. Показать список кинотеатров")
        print("0. Выход")
        print("=" * 60)

        choice = input("Выберите действие: ").strip()

        if choice == "1":
            print("\n--- ДОБАВЛЕНИЕ КИНОТЕАТРА ---")
            name = input("Введите название кинотеатра: ").strip()
            if name:
                system.add_theatre(name)
            else:
                print("Название не может быть пустым!")

        elif choice == "2":
            print("\n--- ДОБАВЛЕНИЕ ЗАЛА ---")
            theatres = system.list_theatres()
            if not theatres:
                print("Сначала добавьте хотя бы один кинотеатр!")
                continue

            print("Доступные кинотеатры:", ", ".join(theatres))
            theatre_name = input("Введите название кинотеатра: ").strip()

            try:
                hall_number = int(input("Введите номер зала: "))
                rows = int(input("Введите количество рядов: "))
                seats = int(input("Введите количество мест в ряду: "))
                system.add_hall(theatre_name, hall_number, rows, seats)
            except ValueError:
                print("Ошибка! Введите числовые значения.")

        elif choice == "3":
            print("\n--- СОЗДАНИЕ СЕАНСА ---")
            theatres = system.list_theatres()
            if not theatres:
                print("Сначала добавьте хотя бы один кинотеатр!")
                continue

            print("Доступные кинотеатры:", ", ".join(theatres))
            theatre_name = input("Введите название кинотеатра: ").strip()

            theatre = system.get_theatre(theatre_name)
            if not theatre:
                print(f"Кинотеатр '{theatre_name}' не найден!")
                continue

            if not theatre["halls"]:
                print("В этом кинотеатре нет залов!")
                continue

            print("Доступные залы:", ", ".join([str(h["number"]) for h in theatre["halls"]]))

            try:
                hall_number = int(input("Введите номер зала: "))
                movie = input("Введите название фильма: ").strip()
                start_time = input("Введите время начала (ГГГГ-ММ-ДД ЧЧ:ММ): ").strip()
                duration = int(input("Введите длительность фильма (в минутах): "))
                system.create_session(theatre_name, hall_number, movie, start_time, duration)
            except ValueError:
                print("Ошибка! Проверьте формат введённых данных.")

        elif choice == "4":
            print("\n--- ПРОДАЖА БИЛЕТА ---")
            theatres = system.list_theatres()
            if not theatres:
                print("Сначала добавьте хотя бы один кинотеатр!")
                continue

            print("Доступные кинотеатры:", ", ".join(theatres))
            theatre_name = input("Введите название кинотеатра: ").strip()

            theatre = system.get_theatre(theatre_name)
            if not theatre:
                print(f"Кинотеатр '{theatre_name}' не найден!")
                continue

            if not theatre["halls"]:
                print("В этом кинотеатре нет залов!")
                continue

            print("Доступные залы:", ", ".join([str(h["number"]) for h in theatre["halls"]]))

            try:
                hall_number = int(input("Введите номер зала: "))

                hall = None
                for h in theatre["halls"]:
                    if h["number"] == hall_number:
                        hall = h
                        break

                if not hall:
                    print(f"Зал №{hall_number} не найден!")
                    continue

                if not hall["sessions"]:
                    print("В этом зале нет сеансов!")
                    continue

                print("\nСеансы:")
                for i, session in enumerate(hall["sessions"]):
                    print(f"{i}. {session['movie']} - {session['start_time']}")

                session_index = int(input("Введите номер сеанса: "))
                system.print_hall_plan(theatre_name, hall_number, session_index)

                row = int(input("Введите номер ряда: ")) - 1
                seat = int(input("Введите номер места: ")) - 1

                system.sell_ticket(theatre_name, hall_number, session_index, row, seat)
            except ValueError:
                print("Ошибка! Введите числовые значения.")

        elif choice == "5":
            print("\n--- ПОИСК БЛИЖАЙШЕГО СЕАНСА ---")
            movie = input("Введите название фильма: ").strip()
            if movie:
                system.find_nearest_session(movie)
            else:
                print("Название фильма не может быть пустым!")

        elif choice == "6":
            print("\n--- ПЛАН ЗАЛА ---")
            theatres = system.list_theatres()
            if not theatres:
                print("Сначала добавьте хотя бы один кинотеатр!")
                continue

            print("Доступные кинотеатры:", ", ".join(theatres))
            theatre_name = input("Введите название кинотеатра: ").strip()

            theatre = system.get_theatre(theatre_name)
            if not theatre:
                print(f"Кинотеатр '{theatre_name}' не найден!")
                continue

            if not theatre["halls"]:
                print("В этом кинотеатре нет залов!")
                continue

            print("Доступные залы:", ", ".join([str(h["number"]) for h in theatre["halls"]]))

            try:
                hall_number = int(input("Введите номер зала: "))

                hall = None
                for h in theatre["halls"]:
                    if h["number"] == hall_number:
                        hall = h
                        break

                if not hall:
                    print(f"Зал №{hall_number} не найден!")
                    continue

                if not hall["sessions"]:
                    print("В этом зале нет сеансов!")
                    continue

                print("\nСеансы:")
                for i, session in enumerate(hall["sessions"]):
                    print(f"{i}. {session['movie']} - {session['start_time']}")

                session_index = int(input("Введите номер сеанса: "))
                system.print_hall_plan(theatre_name, hall_number, session_index)
            except ValueError:
                print("Ошибка! Введите числовые значения.")

        elif choice == "7":
            print("\n--- СПИСОК КИНОТЕАТРОВ ---")
            theatres = system.list_theatres()
            if not theatres:
                print("Кинотеатры не найдены.")
            else:
                for theatre_name in theatres:
                    theatre = system.get_theatre(theatre_name)
                    print(f"\nКинотеатр: {theatre_name}")
                    print(f"Количество залов: {len(theatre['halls'])}")
                    for hall in theatre["halls"]:
                        print(f"  Зал №{hall['number']}: {hall['rows']} рядов x {hall['seats_per_row']} мест, "
                              f"Сеансов: {len(hall['sessions'])}")

        elif choice == "0":
            print("\nСпасибо за использование билетной системы! До свидания!")
            break

        else:
            print("Неверный выбор! Попробуйте снова.")


if __name__ == "__main__":
    main()
