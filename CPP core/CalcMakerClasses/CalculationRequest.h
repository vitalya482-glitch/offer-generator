#ifndef SAM_OFFER_GENERATOR_CALCULATION_REQUEST_H
#define SAM_OFFER_GENERATOR_CALCULATION_REQUEST_H

#include "Position.h"

#include <string>
#include <vector>

// Дополнительный расход, участвующий в расчёте.
struct CalculationExpense {
    // Наименование расхода.
    std::string name;

    // Значение расхода или процент от себестоимости.
    double value = 0.0;

    // Признак того, что значение задано в процентах.
    bool isPercent = false;
};

// Полный входной набор данных для CalcMaker.
struct CalculationRequest {
    // Бренд или направление расчёта.
    std::string brand;

    // Сценарий расчёта: DDP, EXW, Spare Parts и другие.
    std::string scenario;

    // Наименование клиента.
    std::string clientName;

    // Город или пункт поставки.
    std::string deliveryCity;

    // Валюта итогового расчёта.
    std::string currency;

    // Курс для перевода в итоговую валюту.
    double exchangeRate = 1.0;

    // Процент наценки.
    double marginPercent = 0.0;

    // Процент НДС.
    double vatPercent = 0.0;

    // Товарные позиции расчёта.
    std::vector<Position> positions;

    // Дополнительные расходы расчёта.
    std::vector<CalculationExpense> expenses;
};

#endif // SAM_OFFER_GENERATOR_CALCULATION_REQUEST_H
