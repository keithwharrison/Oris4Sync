﻿/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
using System.Collections.Generic;

namespace DotCMIS.Data.Extensions
{
    public interface ICmisExtensionElement
    {
        string Name { get; }

        string Namespace { get; }

        string Value { get; }

        IDictionary<string, string> Attributes { get; }

        IList<ICmisExtensionElement> Children { get; }
    }

    public class CmisExtensionElement :ICmisExtensionElement
    {
        public string Name { get; set; }

        public string Namespace { get; set; }

        public string Value { get; set; }

        public IDictionary<string, string> Attributes { get; set; }

        public IList<ICmisExtensionElement> Children { get; set;}
    }

    public interface IExtensionsData
    {
        IList<ICmisExtensionElement> Extensions { get; set; }
    }

    public class ExtensionsData : IExtensionsData
    {
        public IList<ICmisExtensionElement> Extensions { get; set; }
    }
}
