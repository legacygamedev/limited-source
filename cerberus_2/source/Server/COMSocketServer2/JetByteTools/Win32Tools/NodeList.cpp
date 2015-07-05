///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: NodeList.cpp $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 29/05/02 11:16 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/EchoServerEx/JetByteTools/Win32Tools/NodeList.cpp $
// 
// 2     29/05/02 11:16 Len
// Lint issues.
// 
// 1     24/05/02 12:12 Len
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

#include "NodeList.h"

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Member marked as const indirectly modifies class
//lint -esym(1763, CNodeList::Head) 
//lint -esym(1763, Node::Next) 
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {

///////////////////////////////////////////////////////////////////////////////
// CNodeList
///////////////////////////////////////////////////////////////////////////////

CNodeList::CNodeList() 
   :  m_pHead(0), 
      m_numNodes(0) 
{
}

void CNodeList::PushNode(
   Node *pNode)
{
   pNode->AddToList(this);

   pNode->Next(m_pHead);

   m_pHead = pNode;

   ++m_numNodes;
}

CNodeList::Node *CNodeList::PopNode()
{
   Node *pNode = m_pHead;

   if (pNode)
   {
      RemoveNode(pNode);
   }

   return pNode;
}

CNodeList::Node *CNodeList::Head() const
{
   return m_pHead;
}


size_t CNodeList::Count() const
{
   return m_numNodes;
}

bool CNodeList::Empty() const
{
   return (0 == m_numNodes);
}

void CNodeList::RemoveNode(Node *pNode)
{
   if (pNode == m_pHead)
   {
      //lint -e{613} Possible use of null pointer 
      m_pHead = pNode->Next();
   }

   //lint -e{613} Possible use of null pointer 
   pNode->Unlink();

   --m_numNodes;
}

///////////////////////////////////////////////////////////////////////////////
// CNodeList::Node
///////////////////////////////////////////////////////////////////////////////

CNodeList::Node::Node() 
   :  m_pNext(0), 
      m_pPrev(0), 
      m_pList(0) 
{
}

CNodeList::Node::~Node() 
{
   try
   {
      RemoveFromList();   
   }
   catch(...)
   {
   }

   m_pNext = 0;
   m_pPrev = 0;
   m_pList = 0;
}

CNodeList::Node *CNodeList::Node::Next() const
{
   return m_pNext;
}

void CNodeList::Node::Next(Node *pNext)
{
   m_pNext = pNext;

   if (pNext)
   {
      pNext->m_pPrev = this;
   }
}

void CNodeList::Node::AddToList(CNodeList *pList)
{
   m_pList = pList;
}

void CNodeList::Node::RemoveFromList()
{
   if (m_pList)
   {
      m_pList->RemoveNode(this);
   }
}

void CNodeList::Node::Unlink()
{
   if (m_pPrev)
   {
      m_pPrev->m_pNext = m_pNext;
   }

   if (m_pNext)
   {
      m_pNext->m_pPrev = m_pPrev;
   }
   
   m_pNext = 0;
   m_pPrev = 0;

   m_pList = 0;
}

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools
///////////////////////////////////////////////////////////////////////////////

} // End of namespace JetByteTools 

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -restore
//
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
