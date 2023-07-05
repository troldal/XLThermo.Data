//
// Created by Troldal on 05/07/2023.
//

#ifndef XLTHERMO_DATA_DATABASE_HPP
#define XLTHERMO_DATA_DATABASE_HPP

class Database final
{
public:
    static Database& instance()
    {
        static Database db;    // The one, unique instance
        return db;
    }



private:
    Database() {}                                     // NOLINT
    Database(Database const&)            = delete;    // NOLINT
    Database& operator=(Database const&) = delete;    // NOLINT
    Database(Database&&)                 = delete;    // NOLINT
    Database& operator=(Database&&)      = delete;    // NOLINT
    // ... Potentially some data members
};

#endif    // XLTHERMO_DATA_DATABASE_HPP
