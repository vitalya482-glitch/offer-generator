#ifndef SAM_OFFER_GENERATOR_CORE_H
#define SAM_OFFER_GENERATOR_CORE_H

// Основные библиотеки стандартной C++ библиотеки.
#include <algorithm>
#include <cmath>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <iostream>
#include <map>
#include <memory>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

// Головной манифест C++-ядра приложения.
// Все общие библиотеки и общеиспользуемые определения подключаются отсюда.
// Объявления классов хранятся в отдельных файлах .h, реализации — в .cpp.

#include "CalcMakerClasses/Position.h"
#include "CalcMakerClasses/CalculationRequest.h"
#include "CalcMakerClasses/CalculationResult.h"
#include "CalcMakerClasses/TemplateProfile.h"
#include "CalcMakerClasses/CalcMaker.h"
#include "CalcMakerClasses/PositionParser.h"
#include "CalcMakerClasses/ExcelTemplateRecognizer.h"
#include "CalcMakerClasses/CalcFileCreator.h"
#include "CalcMakerClasses/CalcFileFiller.h"

#endif // SAM_OFFER_GENERATOR_CORE_H
