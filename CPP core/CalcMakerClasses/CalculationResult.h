#ifndef SAM_OFFER_GENERATOR_CALCULATION_RESULT_H
#define SAM_OFFER_GENERATOR_CALCULATION_RESULT_H

#include <string>
#include <vector>

// Итог работы расчётного ядра.
struct CalculationResult {
    // Валюта результата.
    std::string currency;

    // Себестоимость товарных позиций.
    double positionsTotal = 0.0;

    // Общая сумма дополнительных расходов.
    double expensesTotal = 0.0;

    // Цена без наценки.
    double priceWithoutMargin = 0.0;

    // Сумма наценки.
    double marginAmount = 0.0;

    // Цена с наценкой до НДС.
    double priceWithMargin = 0.0;

    // Сумма НДС.
    double vatAmount = 0.0;

    // Финальная стоимость расчёта.
    double grandTotal = 0.0;

    // Предупреждения, которые не остановили расчёт.
    std::vector<std::string> warnings;
};

#endif // SAM_OFFER_GENERATOR_CALCULATION_RESULT_H
