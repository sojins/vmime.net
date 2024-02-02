
#include "stdafx.h"
#include "resource.h"
#include "EmailBuilder.hpp"
#include "Common.hpp"

#pragma managed

namespace vmimeNET {

// A managed wrapper for unmanaged EmailBuilder class

// s_From = "jmiller@gmail.com"  or
// s_From = "John Miller <jmiller@gmail.com>"
EmailBuilder::EmailBuilder(String^ s_From, String^ s_Subject)
{
    try
    {
        mpi_Email = new cEmailBuilder(StrW(s_From), StrW(s_Subject), IDR_MIME_TYPES);
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

// Dispose()
EmailBuilder::~EmailBuilder()
{
    this->!EmailBuilder();
}

// Finalize()
EmailBuilder::!EmailBuilder()
{
    try
    {
        if (mpi_Email) delete mpi_Email;
        mpi_Email = NULL;
    }
    catch (...)
    {
    }
}

void EmailBuilder::SetHeaderField(eHeaderField e_Field, String^ s_Value)
{
    try
    {
        mpi_Email->SetHeaderField((cEmailBuilder::eHeaderField)e_Field, StrW(s_Value));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

// s_Email = "jmiller@gmail.com"  or
// s_Email = "John Miller <jmiller@gmail.com>"
void EmailBuilder::AddTo(String^ s_Email)
{
    try
    {
        mpi_Email->AddTo(StrW(s_Email));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

// s_Email = "jmiller@gmail.com"  or
// s_Email = "John Miller <jmiller@gmail.com>"
void EmailBuilder::AddCc(String^ s_Email)
{
    try
    {
        mpi_Email->AddCc(StrW(s_Email));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

// s_Email = "jmiller@gmail.com"  or
// s_Email = "John Miller <jmiller@gmail.com>"
void EmailBuilder::AddBcc(String^ s_Email)
{
    try
    {
        mpi_Email->AddBcc(StrW(s_Email));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

void EmailBuilder::AddAttachment(String^ s_Path, String^ s_MimeType, String^ s_FileName)
{
    try
    {
        mpi_Email->AddAttachment(StrW(s_Path), StrW(s_MimeType), StrW(s_FileName));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

void EmailBuilder::AddEmbeddedObject(String^ s_Path, String^ s_MimeType, String^ s_CID)
{
    try
    {
        mpi_Email->AddEmbeddedObject(StrW(s_Path), StrW(s_MimeType), StrW(s_CID));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

void EmailBuilder::SetPlainText(String^ s_Plain)
{
    try
    {
        mpi_Email->SetPlainText(StrW(s_Plain));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

void EmailBuilder::SetHtmlText(String^ s_Html)
{
    try
    {
        mpi_Email->SetHtmlText(StrW(s_Html));
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

String^ EmailBuilder::Generate()
{
    try
    {
        return gcnew String(mpi_Email->Generate().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));;
    }
}

} // namespace vmimeNET