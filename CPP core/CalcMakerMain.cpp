#include "Core.h"

#include <fstream>
#include <iostream>
#include <limits>

namespace {

// Разделяет строку по заданному символу.
std::vector<std::string> split(const std::string& value, char separator) {
    std::vector<std::string> parts;
    std::stringstream stream(value);
    std::string part;

    while (std::getline(stream, part, separator)) {
        parts.push_back(part);
    }

    if (!value.empty() && value.back() == separator) {
        parts.emplace_back();
    }

    return parts;
}

// Убирает пробелы по краям строки.
std::string trim(const std::string& value) {
    const std::string whitespace = " \t\r\n";
    const std::size_t first = value.find_first_not_of(whitespace);
    if (first == std::string::npos) {
        return {};
    }

    const std::size_t last = value.find_last_not_of(whitespace);
    return value.substr(first, last - first + 1);
}

// Преобразует строку в число с плавающей точкой.
double parseDouble(const std::string& value, const std::string& fieldName) {
    std::string normalized = trim(value);
    std::replace(normalized.begin(), normalized.end(), ',', '.');

    std::size_t parsedLength = 0;
    const double result = std::stod(normalized, &parsedLength);
    if (parsedLength != normalized.size()) {
        throw std::invalid_argument("Некорректное число в поле: " + fieldName);
    }

    return result;
}

// Преобразует строку в логическое значение.
bool parseBool(const std::string& value) {
    const std::string normalized = trim(value);
    return normalized == "1"
        || normalized == "true"
        || normalized == "yes"
        || normalized == "да";
}

// Загружает запрос расчёта из текстового файла.
CalculationRequest loadRequest(const std::filesystem::path& requestPath) {
    std::ifstream input(requestPath);
    if (!input.is_open()) {
        throw std::runtime_error("Не удалось открыть файл запроса: " + requestPath.string());
    }

    CalculationRequest request;
    std::string line;
    std::size_t lineNumber = 0;

    while (std::getline(input, line)) {
        ++lineNumber;
        line = trim(line);

        if (line.empty() || line.front() == '#') {
            continue;
        }

        const std::size_t separator = line.find('=');
        if (separator == std::string::npos) {
            throw std::invalid_argument("В строке " + std::to_string(lineNumber) + " отсутствует знак '='");
        }

        const std::string key = trim(line.substr(0, separator));
        const std::string value = trim(line.substr(separator + 1));

        if (key == "brand") {
            request.brand = value;
        } else if (key == "scenario") {
            request.scenario = value;
        } else if (key == "client_name") {
            request.clientName = value;
        } else if (key == "delivery_city") {
            request.deliveryCity = value;
        } else if (key == "currency") {
            request.currency = value;
        } else if (key == "exchange_rate") {
            request.exchangeRate = parseDouble(value, key);
        } else if (key == "margin_percent") {
            request.marginPercent = parseDouble(value, key);
        } else if (key == "vat_percent") {
            request.vatPercent = parseDouble(value, key);
        } else if (key == "position") {
            const std::vector<std::string> fields = split(value, '|');
            if (fields.size() < 7) {
                throw std::invalid_argument(
                    "Позиция в строке " + std::to_string(lineNumber)
                    + " должна содержать минимум 7 полей"
                );
            }

            Position position;
            position.code = trim(fields[0]);
            position.name = trim(fields[1]);
            position.brand = trim(fields[2]);
            position.quantity = parseDouble(fields[3], "position.quantity");
            position.unitPrice = parseDouble(fields[4], "position.unit_price");
            position.totalPrice = parseDouble(fields[5], "position.total_price");
            position.currency = trim(fields[6]);
            if (fields.size() > 7) {
                position.source = trim(fields[7]);
            }
            if (fields.size() > 8) {
                position.description = trim(fields[8]);
            }
            request.positions.push_back(position);
        } else if (key == "expense") {
            const std::vector<std::string> fields = split(value, '|');
            if (fields.size() < 3) {
                throw std::invalid_argument(
                    "Расход в строке " + std::to_string(lineNumber)
                    + " должен содержать название, значение и признак процента"
                );
            }

            CalculationExpense expense;
            expense.name = trim(fields[0]);
            expense.value = parseDouble(fields[1], "expense.value");
            expense.isPercent = parseBool(fields[2]);
            request.expenses.push_back(expense);
        } else {
            throw std::invalid_argument("Неизвестный ключ в строке " + std::to_string(lineNumber) + ": " + key);
        }
    }

    return request;
}

// Сохраняет результат расчёта в текстовый файл.
void saveResult(const std::filesystem::path& resultPath, const CalculationResult& result) {
    std::ofstream output(resultPath);
    if (!output.is_open()) {
        throw std::runtime_error("Не удалось создать файл результата: " + resultPath.string());
    }

    output << std::fixed << std::setprecision(2);
    output << "currency=" << result.currency << '\n';
    output << "positions_total=" << result.positionsTotal << '\n';
    output << "expenses_total=" << result.expensesTotal << '\n';
    output << "price_without_margin=" << result.priceWithoutMargin << '\n';
    output << "margin_amount=" << result.marginAmount << '\n';
    output << "price_with_margin=" << result.priceWithMargin << '\n';
    output << "vat_amount=" << result.vatAmount << '\n';
    output << "grand_total=" << result.grandTotal << '\n';

    for (const std::string& warning : result.warnings) {
        output << "warning=" << warning << '\n';
    }
}

// Печатает краткую инструкцию по запуску программы.
void printUsage() {
    std::cout
        << "CalcMaker.exe --self-test\n"
        << "CalcMaker.exe --request request.txt --result result.txt\n";
}

// Выполняет встроенную проверку расчётного ядра.
int runSelfTest() {
    CalculationRequest request;
    request.brand = "Test";
    request.scenario = "EXW";
    request.currency = "EUR";
    request.marginPercent = 20.0;

    Position position;
    position.name = "Test position";
    position.quantity = 1.0;
    position.unitPrice = 100.0;
    position.currency = "EUR";
    request.positions.push_back(position);

    const CalculationResult result = CalcMaker().calculate(request);
    const bool passed = std::abs(result.grandTotal - 120.0) < 0.000001;

    std::cout << (passed ? "CalcMaker self-test: OK\n" : "CalcMaker self-test: FAILED\n");
    return passed ? 0 : 1;
}

} // namespace

// Точка входа консольной версии расчётного ядра.
int main(int argc, char* argv[]) {
    try {
        if (argc == 2 && std::string(argv[1]) == "--self-test") {
            return runSelfTest();
        }

        if (argc == 5 && std::string(argv[1]) == "--request" && std::string(argv[3]) == "--result") {
            const CalculationRequest request = loadRequest(argv[2]);
            const CalculationResult result = CalcMaker().calculate(request);
            saveResult(argv[4], result);
            std::cout << "CalcMaker: расчёт успешно выполнен\n";
            return 0;
        }

        printUsage();
        return 2;
    } catch (const std::exception& error) {
        std::cerr << "CalcMaker: ошибка: " << error.what() << '\n';
        return 1;
    }
}
