
#include "stdafx.h"
#include "resource.h"
#include "EmailParser.hpp"
#include "Common.hpp"

#pragma managed

namespace vmimeNET {

// A managed wrapper for unmanaged EmailParser class

// Do not use this constructor "manually"! 
// Call Pop3.FetchEmailAt() to get an instance of EmailParser instead!
EmailParser::EmailParser(IntPtr p_Internal)
{
    try
    {
        mpi_Email = (cEmailParser*)(void*)p_Internal;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// Dispose()
EmailParser::~EmailParser()
{
    this->!EmailParser();
}

// Finalize()
EmailParser::!EmailParser()
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

// Deletes the email on the server. 
// Connects to the server (POP3 command DELE)
// ATTENTION: If Pop3.Close() is not called afterwards the email(s) will not be deleted!
void EmailParser::Delete()
{
    try
    {
        mpi_Email->Delete();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// Gets the unique identifier that the server has assigned to this email.
// Connects to the server (POP3 command UIDL)
String^ EmailParser::GetUID()
{
    try
    {
        return gcnew String(mpi_Email->GetUID().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// Gets the size of the email in Bytes
// Connects to the server (POP3 command LIST)
UInt32 EmailParser::GetSize()
{
    try
    {
        return mpi_Email->GetSize();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// #############################################################################################
//                                       Message Header
//                  uses only the meail header (fast to download, POP3 uses TOP command)
// #############################################################################################

// Gets the zero-based index of this email in the server's folder (e.g. Inbox)
UInt32 EmailParser::GetIndex()
{
    try
    {
        return mpi_Email->GetIndex();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

EmailParser::eFlags EmailParser::GetFlags()
{
    try
    {
        return (eFlags)mpi_Email->GetFlags();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// For the email "John Miller <jmiller@gmail.com>"
// returns String[] with 
// String[0] = "jmiller@gmail.com" and 
// String[1] = "John Miller"
array<String^>^ EmailParser::GetFrom()
{
    try
    {
        std::wstring s_Name;
        std::wstring s_Email = mpi_Email->GetFrom(&s_Name);

	    array<String^>^ s_Array = gcnew array<String^>(2);
        s_Array[0] = gcnew String(s_Email.c_str());
        s_Array[1] = gcnew String(s_Name .c_str());
        return s_Array;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}
// For the email "John Miller <jmiller@gmail.com>"
// returns String[] with 
// String[0] = "jmiller@gmail.com" and 
// String[1] = "John Miller"
array<String^>^ EmailParser::GetSender()
{
    try
    {
        std::wstring s_Name;
        std::wstring s_Email = mpi_Email->GetSender(&s_Name);

	    array<String^>^ s_Array = gcnew array<String^>(2);
        s_Array[0] = gcnew String(s_Email.c_str());
        s_Array[1] = gcnew String(s_Name .c_str());
        return s_Array;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}
// For the email "John Miller <jmiller@gmail.com>"
// returns String[] with 
// String[0] = "jmiller@gmail.com" and 
// String[1] = "John Miller"
array<String^>^ EmailParser::GetReplyTo()
{
    try
    {
        std::wstring s_Name;
        std::wstring s_Email = mpi_Email->GetReplyTo(&s_Name);

	    array<String^>^ s_Array = gcnew array<String^>(2);
        s_Array[0] = gcnew String(s_Email.c_str());
        s_Array[1] = gcnew String(s_Name .c_str());
        return s_Array;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}


void EmailParser::GetTo(List<String^>^ i_Emails, List<String^>^ i_Names)
{
    try
    {
        std::vector<std::wstring> i_VectMails;
        std::vector<std::wstring> i_VectNames;

        mpi_Email->GetTo(&i_VectMails, &i_VectNames);

        for (size_t i=0; i<i_VectMails.size(); i++)
        {
            i_Emails->Add(gcnew String(i_VectMails.at(i).c_str()));
            i_Names ->Add(gcnew String(i_VectNames.at(i).c_str()));
        }
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}
void EmailParser::GetCc(List<String^>^ i_Emails, List<String^>^ i_Names)
{
    try
    {
        std::vector<std::wstring> i_VectMails;
        std::vector<std::wstring> i_VectNames;

        mpi_Email->GetCc(&i_VectMails, &i_VectNames);

        for (size_t i=0; i<i_VectMails.size(); i++)
        {
            i_Emails->Add(gcnew String(i_VectMails.at(i).c_str()));
            i_Names ->Add(gcnew String(i_VectNames.at(i).c_str()));
        }
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}


String^ EmailParser::GetSubject()
{
    try
    {
        return gcnew String(mpi_Email->GetSubject().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

String^ EmailParser::GetOrganization()
{
    try
    {
        return gcnew String(mpi_Email->GetOrganization().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

String^ EmailParser::GetUserAgent()
{
    try
    {
        return gcnew String(mpi_Email->GetUserAgent().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// returns local time and the timezone in minutes (+/-) deviation from GMT.
// returns DateTime::MinValue (00:00 at January 1, 0001) if no DATE field in the header!
DateTime EmailParser::GetDate([Out] Int32% s32_Timezone)
{
    try
    {
        DateTime i_Date = DateTime::MinValue;
        s32_Timezone  = 0;

        vmime::ref<const vmime::datetime> p_Date = mpi_Email->GetDate();
        if (p_Date)
        {
            i_Date = DateTime(p_Date->getYear(), p_Date->getMonth(),  p_Date->getDay(),
                              p_Date->getHour(), p_Date->getMinute(), p_Date->getSecond());

            s32_Timezone = p_Date->getZone();
        }

        return i_Date;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// #############################################################################################
//                                     Message Body
//    requires the entire message (may be many Megabytes to download, POP3 uses RETR command)
// #############################################################################################

// Gets the entire email (for example to store it in a *.eml file)
String^ EmailParser::GetEmail()
{
    try
    {
        return gcnew String(mpi_Email->GetEmail().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

String^ EmailParser::GetHtmlText()
{
    try
    {
        return gcnew String(mpi_Email->GetHtmlText().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

String^ EmailParser::GetPlainText()
{
    try
    {
        return gcnew String(mpi_Email->GetPlainText().c_str());
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// -----------------------

UInt32 EmailParser::GetEmbeddedObjectCount()
{
    try
    {
        return mpi_Email->GetEmbeddedObjectCount();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

bool EmailParser::GetEmbeddedObjectAt(UInt32 u32_Index, [Out] String^% s_Id, [Out] String^% s_MimeType, [Out] array<Byte>^% u8_Data)
{
    try
    {
        std::wstring id, mimeType;
        std::string  data;
        bool b_Success = mpi_Email->GetEmbeddedObjectAt(u32_Index, id, mimeType, data);

        s_Id       = gcnew String(id.c_str());
        s_MimeType = gcnew String(mimeType.c_str());

        u8_Data = gcnew array<Byte>(data.size());
        Marshal::Copy((IntPtr)(void*)data.c_str(), u8_Data, 0, (int)data.size());

        return b_Success;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

// -----------------------

UInt32 EmailParser::GetAttachmentCount()
{
    try
    {
        return mpi_Email->GetAttachmentCount();
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}

bool EmailParser::GetAttachmentAt(UInt32 u32_Index, [Out] String^% s_Name, [Out] String^% s_MimeType, [Out] array<Byte>^% u8_Data)
{
    try
    {
        std::wstring name, mimeType;
        std::string  data;
        bool b_Success = mpi_Email->GetAttachmentAt(u32_Index, name, mimeType, data);

        s_Name     = gcnew String(name.c_str());
        s_MimeType = gcnew String(mimeType.c_str());

        u8_Data = gcnew array<Byte>(data.size());
        Marshal::Copy((IntPtr)(void*)data.c_str(), u8_Data, 0, (int)data.size());

        return b_Success;
    }
    catch (std::exception& Ex)
    {
        throw gcnew Exception(GCSTRING(Ex.what()));
    }
}


} // namespace vmimeNET