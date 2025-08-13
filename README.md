# 🔧 RevitMEPHoleManager

**Автоматическая расстановка отверстий для MEP элементов в Autodesk Revit**

[![Версия](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/your-username/RevitMEPHoleManager)
[![Revit](https://img.shields.io/badge/Revit-2020%2B-orange.svg)](https://www.autodesk.com/products/revit)
[![.NET](https://img.shields.io/badge/.NET%20Framework-4.8-purple.svg)](https://dotnet.microsoft.com/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

---

## 📖 Описание

**RevitMEPHoleManager** — это мощный плагин для Autodesk Revit, который автоматизирует процесс создания технических отверстий в строительных конструкциях для инженерных коммуникаций (MEP). Плагин анализирует пересечения труб, воздуховодов и кабельных лотков с несущими элементами (стенами, перекрытиями, потолками) и автоматически размещает корректно рассчитанные отверстия.

### 🎯 Основные возможности

- ✅ **Автоматический анализ пересечений** — поиск всех мест, где MEP элементы пересекают конструктивные элементы
- ✅ **Интеллектуальное объединение отверстий** — группировка близко расположенных отверстий в один проем
- ✅ **Точный расчет размеров** — учет диаметров, уклонов и необходимых зазоров
- ✅ **Поддержка сложной геометрии** — работа с наклонными трубами и дуговыми стенами
- ✅ **Исключение проблемных зон** — автоматический пропуск дверных/оконных проемов и несущих элементов
- ✅ **Работа со связанными моделями** — анализ MEP элементов из связанных файлов
- ✅ **Подробная статистика** — отчеты по каждому элементу и хост-конструкции
- ✅ **Лог расчетов** — детальная информация о процессе вычислений

---

## 🚀 Быстрый старт

### Системные требования

- **Autodesk Revit**: 2020, 2021, 2022, 2023, 2024, 2025+
- **.NET Framework**: 4.8 или выше
- **ОС**: Windows 10/11 (64-bit)
- **Память**: минимум 8 ГБ ОЗУ (рекомендуется 16 ГБ)

### Установка

1. **Загрузите** файлы плагина из релизов или соберите из исходного кода
2. **Скопируйте** `RevitMEPHoleManager.dll` в папку плагинов Revit:
   ```
   %APPDATA%\Autodesk\Revit\Addins\2024\
   ```
3. **Создайте** файл манифеста `RevitMEPHoleManager.addin`:
   ```xml
   <?xml version="1.0" encoding="utf-8"?>
   <RevitAddIns>
     <AddIn Type="Application">
       <Name>MEP Hole Manager</Name>
       <Assembly>RevitMEPHoleManager.dll</Assembly>
       <FullClassName>RevitMEPHoleManager.App</FullClassName>
       <ClientId>b83544e2-ec71-4e5f-9739-c23573f0de1c</ClientId>
       <VendorId>ADSK</VendorId>
       <VendorDescription>MEP Tools</VendorDescription>
     </AddIn>
   </RevitAddIns>
   ```
4. **Перезапустите** Revit

### Первый запуск

1. Откройте проект с MEP элементами в Revit
2. Перейдите на вкладку **"MEP Hole Manager"** в ленте
3. Нажмите кнопку **"Запуск"**
4. Выберите семейство отверстий и настройте параметры
5. Нажмите **"Старт"** для анализа и размещения отверстий

---

## 🎛️ Интерфейс пользователя

### Основное окно

![Главное окно плагина](docs/images/main-window.png)

#### 🔧 Панель настроек

| Параметр | Описание | Значение по умолчанию |
|----------|----------|-----------------------|
| **Семейство отверстий** | Generic Model семейство для создания отверстий | Первое доступное |
| **Зазор (мм)** | Дополнительный зазор вокруг MEP элемента | 50 мм |
| **Объединять отверстия** | Включить группировку близких отверстий | Выключено |
| **Макс. расстояние (мм)** | Максимальное расстояние для объединения | 200 мм |

#### 📊 Таблица пересечений

Отображает детальную информацию о каждом найденном пересечении:

- **Хост** — имя конструктивного элемента
- **Host Id** — уникальный идентификатор хоста
- **MEP Id** — идентификатор инженерного элемента
- **Форма** — тип сечения (Round/Rect/Square/Tray)
- **Размеры** — габариты элемента в мм
- **Отверстие** — расчетные размеры отверстия
- **Типоразмер** — имя создаваемого типа семейства
- **Наклон** — есть ли уклон у элемента
- **Расстояние** — зазор до ближайшего элемента

#### 📈 Статистика по хостам

Сводная таблица с количеством отверстий на каждом конструктивном элементе.

#### 📋 Вкладки информации

1. **"Трубы проекта"** — список всех труб с характеристиками
2. **"Лог расчёта"** — подробный журнал вычислений

---

## ⚙️ Техническое описание

### Архитектура плагина

```
RevitMEPHoleManager/
├── 📄 App.cs                    # Точка входа плагина
├── 📄 ShowGuiCommand.cs         # Команда запуска UI
├── 📄 MyMainWindow.xaml(.cs)    # Главное окно интерфейса
├── 📄 IntersectionStats.cs      # Анализ пересечений
├── 📄 Calculaters.cs            # Расчет размеров отверстий
├── 📄 MergeService.cs           # Объединение отверстий
├── 📄 SizeHelper.cs             # Извлечение размеров MEP
├── 📄 HostStatRow.cs            # Модель данных статистики
└── 📄 PipeRow.cs                # Модель данных труб
```

### Поддерживаемые элементы

#### 🏗️ Конструктивные элементы (хосты)

| Тип | Revit Category | Особенности |
|-----|----------------|-------------|
| **Стены** | `OST_Walls` | Прямые и дуговые, многослойные |
| **Перекрытия** | `OST_Floors` | Плоские и наклонные |
| **Потолки** | `OST_Ceilings` | Подвесные и монолитные |

#### 🔧 MEP элементы

| Тип | Revit Category | Параметры размеров |
|-----|----------------|-------------------|
| **Трубы** | `OST_PipeCurves` | Номинальный диаметр (DN) |
| **Воздуховоды** | `OST_DuctCurves` | Ширина × Высота |
| **Кабельные лотки** | `OST_CableTray` | Ширина × Высота |

### Алгоритм работы

#### 1️⃣ Сбор данных

```csharp
// Получение всех хост-элементов
var hosts = new FilteredElementCollector(doc)
    .WherePasses(new LogicalOrFilter(
        new ElementClassFilter(typeof(Wall)),
        new ElementClassFilter(typeof(Floor))))
    .ToElements();

// Сбор MEP из активной модели и связей
var mepElements = CollectMEPFromAllModels(doc);
```

#### 2️⃣ Анализ пересечений

Для каждой пары "хост-MEP":

1. **Получение bounding box** элементов
2. **Проверка пересечения** AABB алгоритмом
3. **Расчет точки пересечения** в центре перекрытия
4. **Определение параметров** элемента (DN, размеры)
5. **Вычисление локальной системы координат** хоста

```csharp
// Пример проверки пересечения
bool Intersects(BoundingBoxXYZ bbox1, BoundingBoxXYZ bbox2)
{
    return (bbox1.Min.X <= bbox2.Max.X && bbox1.Max.X >= bbox2.Min.X) &&
           (bbox1.Min.Y <= bbox2.Max.Y && bbox1.Max.Y >= bbox2.Min.Y) &&
           (bbox1.Min.Z <= bbox2.Max.Z && bbox1.Max.Z >= bbox2.Min.Z);
}
```

#### 3️⃣ Расчет размеров отверстий

**Для прямых элементов:**
```
Размер отверстия = Размер элемента + 2 × Зазор
```

**Для наклонных элементов:**
```
Высота отверстия = Размер элемента / cos(α) + 2 × Зазор
где α — угол между осью элемента и нормалью к стене
```

**Для диагональных труб:**
```
Зазор = 2 × Обычный зазор (удвоенный для компенсации сложной геометрии)
```

#### 4️⃣ Объединение отверстий

Когда включена опция объединения:

1. **Группировка** по хост-элементу
2. **Кластеризация** по расстоянию (если расстояние ≤ порог)
3. **Расчет общего MBR** (минимального охватывающего прямоугольника)
4. **Вычисление размеров кластера:**

```
Ширина кластера = Σ(диаметры/ширины всех элементов) + 2 × общий зазор
Высота кластера = Σ(диаметры/высоты всех элементов) + 2 × общий зазор
```

#### 5️⃣ Фильтрация

**Исключаются отверстия:**
- В дверных и оконных проемах
- Рядом с несущими колоннами и балками  
- При "скользящих" пересечениях (угол < 30°)
- В местах с некорректной геометрией

#### 6️⃣ Размещение отверстий

1. **Выбор подходящей грани** хоста (PickHostFace)
2. **Создание/поиск типоразмера** семейства
3. **Точное позиционирование** с учетом геометрии пересечения
4. **Установка параметров** размеров и уровня
5. **Обработка ошибок** и продолжение при сбоях

---

## 🛠️ Конфигурация

### Настройка семейства отверстий

Плагин работает с Generic Model семействами, имеющими параметры:

| Параметр | Тип | Описание |
|----------|-----|----------|
| `W`, `Width`, `HoleWidth` | Instance | Ширина отверстия |
| `H`, `Height`, `HoleHeight` | Instance | Высота отверстия |
| `Уровень спецификации` | Instance | Уровень для спецификации |
| `Отметка от уровня` | Instance | Высотная отметка |

### Создание семейства отверстий

1. Создайте новое семейство **Generic Model** (face-based)
2. Добавьте параметры экземпляра `Width` и `Height`
3. Создайте 3D геометрию отверстия (void extrusion)
4. Свяжите размеры с параметрами
5. Загрузите семейство в проект

### Настройка манифеста

Для автозагрузки создайте `.addin` файл:

```xml
<?xml version="1.0" encoding="utf-8"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>MEP Hole Manager</Name>
    <Assembly>RevitMEPHoleManager.dll</Assembly>
    <FullClassName>RevitMEPHoleManager.App</FullClassName>
    <ClientId>b83544e2-ec71-4e5f-9739-c23573f0de1c</ClientId>
    <VendorId>ADSK</VendorId>
    <VendorDescription>MEP Tools, v1.0</VendorDescription>
  </AddIn>
</RevitAddIns>
```

---

## 🔬 Примеры использования

### Сценарий 1: Простая труба через стену

**Входные данные:**
- Труба DN150 проходит перпендикулярно через стену 200мм
- Зазор: 25мм

**Результат:**
- Отверстие: 200×200мм (150 + 2×25)
- Типоразмер: "200x200"

### Сценарий 2: Наклонная труба

**Входные данные:**
- Труба DN100 под углом 30° к вертикали
- Зазор: 25мм

**Расчет:**
```
cos(30°) = 0.866
Высота = 100 / 0.866 + 2×25 = 165мм
Ширина = 100 + 2×25 = 150мм
```

**Результат:**
- Отверстие: 150×165мм
- Типоразмер: "150x165"

### Сценарий 3: Кластер труб

**Входные данные:**
- 3 трубы: DN100, DN150, DN100
- Расстояние между осями: 200мм, 200мм
- Зазор: 25мм, объединение: 300мм

**Расчет:**
```
Ширина = 100 + 150 + 100 + 2×25 = 400мм
Высота = 100 + 150 + 100 + 2×25 = 400мм
```

**Результат:**
- Одно отверстие: 400×400мм
- Типоразмер: "400x400"

---

## 🐛 Диагностика и устранение неисправностей

### Часто встречающиеся проблемы

#### ❌ "Семейство не найдено"

**Причина:** В проекте нет Generic Model семейств  
**Решение:** Загрузите face-based семейство отверстий

#### ❌ "Нет пересечений"

**Причины:**
- MEP элементы не пересекают конструктивы
- Элементы находятся в проемах окон/дверей
- Некорректная геометрия модели

**Решение:** Проверьте модель в 3D виде

#### ❌ "Отверстия не создаются"

**Причины:**
- Отсутствуют права на редактирование
- Элементы заблокированы в рабочих наборах
- Ошибки в семействе отверстий

**Решение:** 
- Проверьте права доступа
- Разблокируйте рабочие наборы
- Проверьте параметры семейства

#### ❌ ArgumentNullException

**Причина:** Некорректная геометрия элементов  
**Решение:** Плагин автоматически пропускает проблемные элементы

### Производительность

#### Для больших моделей:

- **Отключите** объединение отверстий для первого прогона
- **Ограничьте** видимость элементов в текущем виде
- **Закройте** ненужные связанные модели
- **Используйте** рабочие наборы для фильтрации

#### Оптимальные настройки:

| Размер модели | Зазор | Объединение | Макс. расстояние |
|---------------|-------|-------------|------------------|
| < 1000 элементов | 25мм | ✅ | 200мм |
| 1000-5000 | 25мм | ✅ | 150мм |
| > 5000 | 50мм | ❌ | — |

---

## 🔧 Сборка из исходного кода

### Требования для разработки

- **Visual Studio 2019+** с поддержкой .NET Framework 4.8
- **Revit 2024 SDK** (автоматически из установки Revit)
- **Git** для клонирования репозитория

### Шаги сборки

1. **Клонирование репозитория:**
```bash
git clone https://github.com/your-username/RevitMEPHoleManager.git
cd RevitMEPHoleManager
```

2. **Открытие в Visual Studio:**
```bash
start RevitMEPHoleManager.sln
```

3. **Настройка путей к API Revit:**
   - Откройте файл проекта `.csproj`
   - Измените пути к `RevitAPI.dll` и `RevitAPIUI.dll` если нужно:
   ```xml
   <Reference Include="RevitAPI">
     <HintPath>C:\Program Files\Autodesk\Revit 2024\RevitAPI.dll</HintPath>
   </Reference>
   ```

4. **Сборка проекта:**
   - Выберите конфигурацию **Release**
   - Нажмите **Build → Build Solution** (Ctrl+Shift+B)

5. **Результат:**
   - Файл `bin\Release\RevitMEPHoleManager.dll`
   - Готов для установки в Revit

### Настройка разработки

**Автокопирование в папку Revit:**
```xml
<Target Name="PostBuild" AfterTargets="PostBuildEvent">
  <Copy SourceFiles="$(TargetPath)" 
        DestinationFolder="%APPDATA%\Autodesk\Revit\Addins\2024\" />
</Target>
```

**Отладка:**
1. Установите Revit как стартовое приложение
2. Включите "Debug → Attach to Process → Revit.exe"
3. Поставьте breakpoints в коде
4. Запустите плагин в Revit

---

## 📚 API Reference

### Основные классы

#### IntersectionStats
Анализ пересечений MEP элементов с конструктивами.

```csharp
public static (int wRnd, int wRec, int fRnd, int fRec, 
               List<IntersectRow> rows, List<HostStatRow> hostStats) 
       Analyze(IEnumerable<Element> hosts, 
               IEnumerable<(Element elem, Transform tx)> mepList, 
               double clearanceMm)
```

#### Calculaters
Расчет размеров отверстий с учетом уклонов.

```csharp
public static void GetHoleSizeIncline(bool isRound, double elemWmm, 
                                      double elemHmm, double clearanceMm, 
                                      XYZ axisLocal, out double holeWmm, 
                                      out double holeHmm)
```

#### MergeService
Объединение близко расположенных отверстий.

```csharp
public static IEnumerable<IntersectRow> Merge(IEnumerable<IntersectRow> rows, 
                                              double gapMm, double clearanceMm, 
                                              HoleLogger log)
```

#### SizeHelper
Извлечение размеров MEP элементов.

```csharp
public static bool TryGetSizes(Element mep, out double wMm, 
                               out double hMm, out ShapeKind shape)
```

### Модели данных

#### IntersectRow
```csharp
public class IntersectRow 
{
    public int HostId { get; set; }
    public int MepId { get; set; }
    public XYZ Center { get; set; }
    public XYZ PipeDir { get; set; }
    public XYZ LocalCtr { get; set; }
    public XYZ AxisLocal { get; set; }
    public double WidthLocFt { get; set; }
    public double HeightLocFt { get; set; }
    public double HoleWidthMm { get; set; }
    public double HoleHeightMm { get; set; }
    public string HoleTypeName { get; set; }
    public bool IsMerged { get; set; }
    public bool IsDiagonal { get; set; }
    public XYZ GroupCtr { get; set; }
    public double? GapMm { get; set; }
}
```

#### HostStatRow
```csharp
public class HostStatRow 
{
    public int HostId { get; set; }
    public string HostName { get; set; }
    public int Round { get; set; }
    public int Rect { get; set; }
}
```

---

## 🤝 Участие в разработке

Мы приветствуем вклад в развитие проекта!

### Как внести свой вклад

1. **Fork** репозитория
2. **Создайте** ветку для новой функции (`git checkout -b feature/amazing-feature`)
3. **Внесите** изменения и **протестируйте** их
4. **Зафиксируйте** изменения (`git commit -m 'Add amazing feature'`)
5. **Отправьте** в ветку (`git push origin feature/amazing-feature`)
6. **Создайте** Pull Request

### Стандарты кодирования

- **C# Naming Conventions** по Microsoft
- **XML Documentation** для публичных методов
- **Unit Tests** для новой функциональности
- **Error Handling** с try-catch блоками
- **Performance** — избегайте лишних итераций по элементам

### Приоритетные направления

- 🔄 Поддержка дополнительных MEP категорий
- 🎯 Улучшение алгоритмов позиционирования
- 📊 Расширенная отчетность и экспорт
- 🌐 Локализация интерфейса
- ⚡ Оптимизация производительности
- 🔌 Интеграция с BIM360/ACC

---

## 📄 Лицензия

Этот проект распространяется под лицензией MIT. Подробности в файле [LICENSE](LICENSE).

```
MIT License

Copyright (c) 2025 RevitMEPHoleManager

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👥 Авторы и благодарности

- **Основной разработчик**: [Ваше имя](https://github.com/your-username)
- **Архитектура**: Модульная система с разделением ответственности
- **Тестирование**: Команда BIM-инженеров
- **Документация**: Техническая команда

### Особая благодарность

- **Autodesk Revit API** за мощный инструментарий разработки
- **Сообществу разработчиков Revit** за советы и best practices
- **Beta-тестерам** за обратную связь и выявление ошибок

---

## 📞 Поддержка и контакты

### Получение помощи

- 📧 **Email**: support@revittools.com
- 💬 **Telegram**: [@revitmetools](https://t.me/revitmetools)
- 🐛 **Issues**: [GitHub Issues](https://github.com/your-username/RevitMEPHoleManager/issues)
- 📖 **Wiki**: [Документация](https://github.com/your-username/RevitMEPHoleManager/wiki)

### Отчеты об ошибках

При создании issue укажите:

1. **Версию Revit** и **версию плагина**
2. **Описание проблемы** и шаги для воспроизведения
3. **Скриншоты** интерфейса и ошибок
4. **Лог-файлы** из вкладки "Лог расчёта"
5. **Тестовую модель** (если возможно)

### Запросы на улучшения

Мы открыты к предложениям! Создавайте issues с тегом `enhancement`.

---

## 🔄 История изменений

### v1.0.0 (2025-01-XX)

#### ✨ Новые функции
- Автоматический анализ пересечений MEP элементов
- Интеллектуальное объединение отверстий
- Поддержка наклонных труб и дуговых стен
- Исключение проемов и несущих элементов
- Работа со связанными моделями
- Подробная статистика и логирование

#### 🔧 Технические улучшения
- Модульная архитектура для легкого расширения
- Защита от ошибок геометрии (ArgumentNullException)
- Оптимизированные алгоритмы расчета
- Совместимость с Revit 2020-2025+

#### 🐛 Исправленные ошибки
- Обработка null references в геометрии
- Корректная работа с CylindricalFace
- Стабильное позиционирование отверстий
- Правильный расчет зазоров для кластеров

---

## 🎯 Дорожная карта

### v1.1.0 (планируется Q2 2025)
- 🔧 Поддержка кабельных каналов и трубопроводов
- 📊 Экспорт отчетов в Excel/PDF
- 🎨 Настраиваемые стили и шаблоны отверстий
- 🔄 Batch processing для множественных файлов

### v1.2.0 (планируется Q3 2025)
- 🌐 Английская локализация
- 🔌 API для сторонних разработчиков
- 📱 Мобильная версия отчетов
- ☁️ Интеграция с облачными сервисами

### v2.0.0 (планируется Q4 2025)
- 🤖 Машинное обучение для оптимизации размещения
- 🏗️ Поддержка арматуры и закладных деталей
- 📐 3D визуализация планируемых отверстий
- 🔄 Синхронизация между дисциплинами

---

**⭐ Если плагин оказался полезным, поставьте звезду на GitHub!**

**📢 Следите за обновлениями и делитесь опытом использования!**

---

*Последнее обновление: 2025-01-XX*