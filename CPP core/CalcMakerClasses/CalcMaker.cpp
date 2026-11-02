#include "CalcMaker.h"

#include <cmath>
#include <stdexcept>

// Создаёт объект расчётного ядра.
CalcMaker::CalcMaker() = default;

// Освобождает ресурсы объекта расчётного ядра.
CalcMaker::~CalcMaker() = default;

// Выполняет расчёт по входному запросу.
CalculationResult CalcMaker::calculate(const CalculationRequest& request) const {
    validate(request);

    CalculationResult result;
    result.currency = request.currency;

    for (const Position& position : request.positions) {
        const double positionTotal = position.totalPrice > 0.0
            ? position.totalPrice
            : position.quantity * position.unitPrice;
        result.positionsTotal += positionTotal;
    }

    for (const CalculationExpense& expense : request.expenses) {
        const double expenseAmount = expense.isPercent
            ? result.positionsTotal * expense.value / 100.0
            : expense.value;
        result.expensesTotal += expenseAmount;
    }

    result.priceWithoutMargin = result.positionsTotal + result.expensesTotal;
    result.marginAmount = result.priceWithoutMargin * request.marginPercent / 100.0;
    result.priceWithMargin = result.priceWithoutMargin + result.marginAmount;
    result.vatAmount = result.priceWithMargin * request.vatPercent / 100.0;
    result.grandTotal = result.priceWithMargin + result.vatAmount;

    return result;
}

// Проверяет входные данные расчёта.
void CalcMaker::validate(const CalculationRequest& request) const {
    if (request.positions.empty()) {
        throw std::invalid_argument("В расчёте должна быть хотя бы одна позиция");
    }

    if (request.currency.empty()) {
        throw std::invalid_argument("Не указана валюта расчёта");
    }

    if (!std::isfinite(request.exchangeRate) || request.exchangeRate <= 0.0) {
        throw std::invalid_argument("Курс должен быть положительным числом");
    }

    if (!std::isfinite(request.marginPercent) || request.marginPercent < 0.0) {
        throw std::invalid_argument("Наценка не может быть отрицательной");
    }

    if (!std::isfinite(request.vatPercent) || request.vatPercent < 0.0) {
        throw std::invalid_argument("НДС не может быть отрицательным");
    }

    for (const Position& position : request.positions) {
        if (position.name.empty()) {
            throw std::invalid_argument("У позиции не указано наименование");
        }
        if (!std::isfinite(position.quantity) || position.quantity <= 0.0) {
            throw std::invalid_argument("Количество позиции должно быть положительным");
        }
        if (!std::isfinite(position.unitPrice) || position.unitPrice < 0.0) {
            throw std::invalid_argument("Цена позиции не может быть отрицательной");
        }
        if (!position.currency.empty() && position.currency != request.currency) {
            throw std::invalid_argument("В расчёте обнаружены разные валюты");
        }
    }
}

// Сбрасывает внутреннее состояние расчётного ядра.
void CalcMaker::reset() {
    // На текущем этапе CalcMaker не хранит изменяемого состояния.
}
