#include "CalcFileCreator.h"

#include <stdexcept>

// Создаёт объект генератора файлов расчёта.
CalcFileCreator::CalcFileCreator() = default;

// Копирует шаблон в папку результата и возвращает путь нового файла.
std::filesystem::path CalcFileCreator::create(
    const std::filesystem::path& templatePath,
    const std::filesystem::path& outputDirectory,
    const std::string& fileName
) const {
    if (templatePath.empty() || !std::filesystem::exists(templatePath)) {
        throw std::runtime_error("Шаблон расчёта не найден");
    }

    if (outputDirectory.empty()) {
        throw std::invalid_argument("Не указана папка результата");
    }

    if (fileName.empty()) {
        throw std::invalid_argument("Не указано имя нового файла расчёта");
    }

    std::filesystem::create_directories(outputDirectory);
    const std::filesystem::path outputPath = outputDirectory / fileName;

    if (std::filesystem::exists(outputPath)) {
        throw std::runtime_error("Файл расчёта уже существует: " + outputPath.string());
    }

    std::filesystem::copy_file(
        templatePath,
        outputPath
    );
    return outputPath;
}
