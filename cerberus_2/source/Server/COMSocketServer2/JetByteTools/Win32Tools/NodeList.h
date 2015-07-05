#if defined (_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

#ifndef JETBYTE_TOOLS_NODE_LIST_INCLUDED__
#define JETBYTE_TOOLS_NODE_LIST_INCLUDED__
///////////////////////////////////////////////////////////////////////////////
//
// File           : $Workfile: NodeList.h $
// Version        : $Revision: 2 $
// Function       : 
//
// Author         : $Author: Len $
// Date           : $Date: 29/05/02 11:31 $
//
// Notes          : 
//
// Modifications  :
//
// $Log: /Web Articles/SocketServers/EchoServerEx/JetByteTools/Win32Tools/NodeList.h $
// 
// 2     29/05/02 11:31 Len
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

///////////////////////////////////////////////////////////////////////////////
// Lint options
//
//lint -save
//
// Member hides non virtual member
//lint -esym(1511, CNodeList::Head)
//lint -esym(1511, CNodeList::PopNode)
//lint -esym(1511, CNodeList::PushNode)
//
///////////////////////////////////////////////////////////////////////////////

#include <wtypes.h>

///////////////////////////////////////////////////////////////////////////////
// Namespace: JetByteTools
///////////////////////////////////////////////////////////////////////////////

namespace JetByteTools {

///////////////////////////////////////////////////////////////////////////////
// CNodeList
///////////////////////////////////////////////////////////////////////////////

class CNodeList
{
   public :

      class Node
      {
         public :

            Node *Next() const;

            void Next(Node *pNext);

            void AddToList(CNodeList *pList);

            void RemoveFromList();

         protected :

            Node();
            ~Node();

         private :

            friend class CNodeList;

            void Unlink();

            Node *m_pNext;
            Node *m_pPrev;

            CNodeList *m_pList;
      };

      CNodeList();

      void PushNode(Node *pNode);

      Node *PopNode();

      Node *Head() const;

      size_t Count() const;

      bool Empty() const;

   private :

      friend void Node::RemoveFromList();

      void RemoveNode(Node *pNode);

      Node *m_pHead; 

      size_t m_numNodes;
};

///////////////////////////////////////////////////////////////////////////////
// TNodeList
///////////////////////////////////////////////////////////////////////////////

template <class T> class TNodeList : public CNodeList
{
   public :
   
      void PushNode(T *pNode);
      
      T *PopNode();
   
      T *Head() const;

      static T *Next(const T *pNode);
};

template <class T>
void TNodeList<T>::PushNode(T *pNode)
{
   CNodeList::PushNode(pNode);
}

template <class T>
T *TNodeList<T>::PopNode()
{
   return static_cast<T*>(CNodeList::PopNode());
}

template <class T>
T *TNodeList<T>::Head() const
{
   return static_cast<T*>(CNodeList::Head());
}

template <class T>
T *TNodeList<T>::Next(const T *pNode)
{
   return static_cast<T*>(pNode->Next());
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

#endif //JETBYTE_TOOLS_NODE_LIST_INCLUDED__

///////////////////////////////////////////////////////////////////////////////
// End of file
///////////////////////////////////////////////////////////////////////////////
