#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_WIN32_IO_BUFFER_INCLUDED__
#define JETBYTE_TOOLS_WIN32_IO_BUFFER_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: IOBuffer.h $
// Version        : $Revision: 11 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 7/06/02 14:14 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/SimpleProtocolServer2/JetByteTools/Win32Tools/IOBuffer.h $
// 
// 11    7/06/02 14:14 Len
// We now derive from OVERLAPPED rather than contain an instance. This
// means we can remove the explicit conversion functions.
// 
// 10    29/05/02 12:05 Len
// More lint issues.
// 
// 9     29/05/02 11:35 Len
// Lint issues.
// 
// 8     26/05/02 15:11 Len
// Factored out common 'user data' code into a mixin base class.
// Use NodeList for the list implementations. 
// 
// 7     21/05/02 11:35 Len
// User data can now be stored/retrieved as either an unsigned long or a
// void *.
// 
// 6     20/05/02 23:17 Len
// Updated copyright and disclaimers.
// 
// 5     20/05/02 8:07 Len
// Refactored to remove the concept of the io operation that the buffer is
// being used for and replace it with a more abstract 'user data' concept.
// The knowledge of what the buffer is used for is now in the SocketServer
// code and users are free to add their own user data if they use the
// buffers outside of the SocketServer.
// General code cleanup.
// 
// 4     14/05/02 13:49 Len
// Our optimised SplitData() method is inappropriate now that we're
// reference counting the buffers. We were allocating a new buffer and
// putting the extra data into it. We should have been putting the data we
// wanted to process into the new buffer and then moving the remaining
// data in the existing buffer to close the gap. The new code results in
// one more memory move.
// 
// 3     13/05/02 13:45 Len
// Removed all knowledge of sockets from the iobuffer.
// Added the concept of 'max free buffers' to the allocator. This allows
// us to shrink the buffer pool rather than just growing it. IOBuffers are
// now reference counted.
// 
// 2     10/05/02 19:25 Len
// Lint options and code cleaning.
// 
// 1     9/05/02 18:47 Len
// 
///////////////////////////////////////////////////////////////////////////////
//
// Copyright 1997 - 2002 JetByte Limited.
//
// JetByte Limited grants you ("Licensee") a non-exclusive, royalty free, 
// licence to use, modify and redistribute this software in source and binary 
// code form, provided that i) this copyright notice and licence appear on all 
// copies of the software; and ii) Licensee does not utilize the software in a 
// manner which is disparaging to JetByte Limited.
//
// This software is provided "as is" without a warranty of any kind. All 
// express or implied conditions, representations and warranties, including
// any implied warranty of merchantability, fitness for a particular purpose
// or non-infringement, are hereby excluded. JetByte Limited and its licensors 
// shall not be liable for any damages suffered by licensee as a result of 
// using, modifying or distributing the software or its derivatives. In no
// event will JetByte Limited be liable for any lost revenue, profit or data,
// or for direct, indirect, special, consequential, incidental or punitive
// damages, however caused and regardless of the theory of liability, arising 
// out of the use of or inability to use software, even if JetByte Limited 
// has been advised of the possibility of such damages.
//
// This software is not designed or intended for use in on-line control of 
// aircraft, air traffic, aircraft navigation or aircraft communications; or in 
// the design, construction, operation or maintenance of any nuclear 
// facility. Licensee represents and warrants that it will not use or 
// redistribute the Software for such purposes. 
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//lint -esym(1725, CIOBuffer::m_allocator)   class member is a reference     
// 
// Private constructor
//lint -esym(1704, CIOBuffer::CIOBuffer)
//lint -esym(1704, Allocator::Allocator)
//
// No default constructor
//lint -esym(1712, CIOBuffer)
//lint -esym(1712, Allocator)
//
// Base class destructor not virtual
//lint -esym(1509, Node)
//lint -esym(1509, COpaqueUserData)
//
// Base class has no destructor
//lint -esym(1510, _OVERLAPPED)
//
// Symbol hides global operator
//lint -esym(1737, CIOBuffer::operator new)
//
// 
//lint -e537 repeated include winsock2.h
//
///////////////////////////////////////////////////////////////////////////////

#include <winsock2.h>

#include "CriticalSection.h" 
#include "tstring.h"

#include "NodeList.h"
#include "OpaqueUserData.h"

#pragma warning(disable: 4200) // nonstandard extension used : zero-sized array in struct/union

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {
namespace Win32 {

///////////////////////////////////////////////////////////////////////////////
// CIOBuffer::Allocator
///////////////////////////////////////////////////////////////////////////////

class CIOBuffer : public OVERLAPPED, public CNodeList::Node, public COpaqueUserData
{
   public :

      class Allocator;

      friend class Allocator;

      WSABUF *GetWSABUF() const;

      size_t GetUsed() const;

      size_t GetSize() const;

      const BYTE *GetBuffer() const;

      void SetupRead();

      void SetupWrite();
      
      void AddData(
         const char * const pData,
         size_t dataLength);

      void AddData(
         const BYTE * const pData,
         size_t dataLength);

      void AddData(
         BYTE data);

      void Use(
         size_t dataUsed);

      CIOBuffer *SplitBuffer(
         size_t bytesToRemove);

      CIOBuffer *AllocateNewBuffer() const;

      void ConsumeAndRemove(
         size_t bytesToRemove);

      void Empty();

      void AddRef();
      void Release();

      size_t GetOperation() const;
      
      void SetOperation(size_t operation);

   private :

      size_t m_operation;

      WSABUF m_wsabuf;

      Allocator &m_allocator;

      long m_ref;
      const size_t m_size;
      size_t m_used;
      BYTE m_buffer[0];      // start of the actual buffer, must remain the last
                             // data member in the class.

   private :

      static void *operator new(size_t objSize, size_t bufferSize);
      static void operator delete(void *pObj, size_t bufferSize);
      
      CIOBuffer(
         Allocator &m_allocator,
         size_t size);

      // No copies do not implement
      CIOBuffer(const CIOBuffer &rhs);
      CIOBuffer &operator=(const CIOBuffer &rhs);
};

///////////////////////////////////////////////////////////////////////////////
// CIOBuffer::Allocator
///////////////////////////////////////////////////////////////////////////////

class CIOBuffer::Allocator
{
   public :

      friend class CIOBuffer;

      explicit Allocator(
         size_t bufferSize,
         size_t maxFreeBuffers);

      virtual ~Allocator();

      CIOBuffer *Allocate();

   protected :

      void Flush();

   private :

      void Release(
         CIOBuffer *pBuffer);

      virtual void OnBufferCreated() {}
      virtual void OnBufferAllocated() {}
      virtual void OnBufferReleased() {}
      virtual void OnBufferDestroyed() {}

      void DestroyBuffer(
         CIOBuffer *pBuffer);

      const size_t m_bufferSize;

      typedef TNodeList<CIOBuffer> BufferList;
      
      BufferList m_freeList;
      BufferList m_activeList;
      
      const size_t m_maxFreeBuffers;

      CCriticalSection m_criticalSection;

      // No copies do not implement
      Allocator(const Allocator &rhs);
      Allocator &operator=(const Allocator &rhs);
};

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools::Win32
///////////////////////////////////////////////////////////////////////////////

} // End of namespace Win32
} // End of namespace JetByteTools 

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

#endif // JETBYTE_TOOLS_WIN32_IO_BUFFER_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
